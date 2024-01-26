VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "selreporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   11115
   Begin Threed.SSPanel panToolBar 
      Height          =   4980
      Index           =   0
      Left            =   10380
      TabIndex        =   5
      Top             =   810
      Width           =   690
      _Version        =   65536
      _ExtentX        =   1217
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
         TabIndex        =   13
         Top             =   15
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
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
         Left            =   120
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1050
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
         Picture         =   "selreporte.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1470
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
         Picture         =   "selreporte.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2235
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
         Picture         =   "selreporte.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2670
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
         Picture         =   "selreporte.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3825
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
         Picture         =   "selreporte.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Tag             =   "0"
         Top             =   4260
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
         Picture         =   "selreporte.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "0"
         Top             =   615
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
         Picture         =   "selreporte.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3090
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
         Picture         =   "selreporte.frx":00D0
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   4980
      Left            =   5460
      TabIndex        =   14
      Top             =   810
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   8784
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   494
      TabMaxWidth     =   2205
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
      TabPicture(0)   =   "selreporte.frx":00EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dcaSeleccion(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbSeleccion(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Entidad Pensión"
      TabPicture(1)   =   "selreporte.frx":0108
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dcaSeleccion(1)"
      Tab(1).Control(1)=   "tdbSeleccion(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ubicación"
      TabPicture(2)   =   "selreporte.frx":0124
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dcaSeleccion(2)"
      Tab(2).Control(1)=   "tdbSeleccion(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Sección"
      TabPicture(3)   =   "selreporte.frx":0140
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dcaSeleccion(3)"
      Tab(3).Control(1)=   "tdbSeleccion(3)"
      Tab(3).ControlCount=   2
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4155
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   90
         Width           =   4670
         _ExtentX        =   8229
         _ExtentY        =   7329
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
         Left            =   90
         Top             =   4260
         Width           =   4670
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
         Height          =   4155
         Index           =   1
         Left            =   -74910
         TabIndex        =   16
         Top             =   90
         Width           =   4670
         _ExtentX        =   8229
         _ExtentY        =   7329
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
         Left            =   -74910
         Top             =   4260
         Width           =   4670
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
         Height          =   4155
         Index           =   2
         Left            =   -74910
         TabIndex        =   24
         Top             =   90
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   7329
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
         Left            =   -74910
         Top             =   4260
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
         Height          =   4155
         Index           =   3
         Left            =   -74910
         TabIndex        =   25
         Top             =   90
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   7329
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
         Index           =   3
         Left            =   -74910
         Top             =   4260
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
   Begin Threed.SSFrame frmCuadro 
      Height          =   4455
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   1305
      Width           =   5385
      _Version        =   65536
      _ExtentX        =   9499
      _ExtentY        =   7858
      _StockProps     =   14
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
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   3975
         Index           =   4
         Left            =   30
         TabIndex        =   23
         Top             =   105
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   7011
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
         Index           =   4
         Left            =   30
         Top             =   4110
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
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1035
      TabIndex        =   18
      Top             =   3000
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
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11115
      _Version        =   65536
      _ExtentX        =   19606
      _ExtentY        =   1296
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
         Left            =   3495
         MaxLength       =   4
         TabIndex        =   2
         Top             =   225
         Width           =   705
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   10005
         TabIndex        =   21
         Top             =   75
         Visible         =   0   'False
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
         PictureUp       =   "selreporte.frx":015C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   9600
         TabIndex        =   22
         Top             =   75
         Visible         =   0   'False
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
         PictureUp       =   "selreporte.frx":0178
      End
      Begin Threed.SSRibbon ribFormato 
         Height          =   360
         Index           =   1
         Left            =   915
         TabIndex        =   26
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
         PictureUp       =   "selreporte.frx":0194
      End
      Begin Threed.SSRibbon ribFormato 
         Height          =   360
         Index           =   0
         Left            =   510
         TabIndex        =   27
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
         PictureUp       =   "selreporte.frx":01B0
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   4290
         TabIndex        =   19
         Top             =   225
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
         Left            =   4695
         TabIndex        =   3
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Formato Reporte :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   2130
         TabIndex        =   1
         Top             =   270
         Width           =   1320
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   585
         Index           =   0
         Left            =   1905
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   7110
      End
   End
   Begin Threed.SSCheck chkPeriodos 
      Height          =   225
      Left            =   135
      TabIndex        =   20
      Top             =   945
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Analisis Periodos"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
   End
End
Attribute VB_Name = "fSelReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(5, 2)                            ' Array de inicio y fin de seleccion de registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Private Sub GeneraReporte(nTabIndex As Integer, s_Tabla As String, s_Reporte As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim a_Columnas(), a_Totales() As Double, a_Niveles() As Double
  Dim nNivel As Integer, sValor As String, nValor As Double
  Dim sCabecera As String, sDetalle As String, sTotal As String
  Dim a_Detalle(), sPersonal As String, nSecuencia As Long
  Dim a_Quiebre() As Double, sQuiebre As String, sDesquiebre As String
  Dim sSeccion As String
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Registros detalle con campos
  s_Sql = "SELECT rpt.alias, rpt.descripcion, rpt.orden, rpt.tipo, "
  s_Sql = s_Sql & "rpt.nivel, rpt.signo, rpt.impreso, rpt.longitud, vfx.valor "
  s_Sql = s_Sql & "FROM pldetareporte rpt "
  s_Sql = s_Sql & "LEFT JOIN plvarfunc vfx ON rpt.tipo=vfx.tipo AND rpt.alias=vfx.codigo AND IFNULL(vfx.valor, '')<>'' "
  s_Sql = s_Sql & "WHERE rpt.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND rpt.codrpt='" & s_Reporte & "' "
  s_Sql = s_Sql & "ORDER BY rpt.orden"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    n_Index = 0: s_Sql = ""
    nNivel = n_Index
    While Not porstRecordset.EOF
      n_Index = n_Index + 1
      ' Redimensiono el arreglo de las cabeceras
      ReDim Preserve a_Columnas(8, n_Index)
      a_Columnas(1, n_Index) = Trim(porstRecordset("alias"))
      a_Columnas(7, n_Index) = CInt(porstRecordset("longitud"))
      a_Columnas(2, n_Index) = UCase(Left(Trim(porstRecordset("descripcion")), a_Columnas(7, n_Index)))
      a_Columnas(3, n_Index) = Trim(porstRecordset("tipo"))
      a_Columnas(4, n_Index) = Trim(porstRecordset("nivel"))
      a_Columnas(5, n_Index) = Trim(porstRecordset("signo"))
      a_Columnas(6, n_Index) = Trim(porstRecordset("impreso"))
      a_Columnas(8, n_Index) = gdl_Funcion.aTexto(porstRecordset("valor"))
      If a_Columnas(8, n_Index) <> "" And a_Columnas(8, n_Index) <> "NULL" Then
        s_Sql = s_Sql & Trim(porstRecordset!valor) & " AS dato_" & Format(n_Index, "00") & ", "
      End If
      nNivel = CInt(IIf(CInt(porstRecordset("nivel")) > nNivel, porstRecordset("nivel"), nNivel))
      If a_Columnas(3, n_Index) = "D" Then
        sCabecera = sCabecera & gdl_Funcion.PadR(a_Columnas(2, n_Index), a_Columnas(7, n_Index), Chr(32))
      Else
        sCabecera = sCabecera & gdl_Funcion.PadL(a_Columnas(2, n_Index), a_Columnas(7, n_Index), Chr(32))
      End If
      porstRecordset.MoveNext
    Wend
  End If
  
  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT " & s_Sql
  s_Sql = s_Sql & "rpt.codrpt, rpt.alias, rpt.descripcion, rpt.orden, rpt.tipo, rpt.nivel, rpt.signo, rpt.impreso, rpt.longitud, "
  s_Sql = s_Sql & "res.codpsn, ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)), 2) AS nimporte, "
  s_Sql = s_Sql & "dxr.codcco, cco.detcco, dxr.codafp, afp.desafp, dxr.codubica, ubi.desubica, dxr.codsec, sec.dessec "
  s_Sql = s_Sql & "FROM pldetareporte rpt "
  s_Sql = s_Sql & "INNER JOIN plresultado res ON rpt.codcls=res.codcls AND rpt.alias=res.codcpc AND rpt.tipo='C' "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON dxr.codeps=eps.codeps "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE rpt.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND rpt.codrpt='" & s_Reporte & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  If nTabIndex <> 4 Then
    s_Sql = s_Sql & "AND dxr." & Choose(nTabIndex + 1, "codcco", "codafp", "codubica", "codsec") & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "codcco, codpsn, codcpc ", "codafp, codpsn, codcpc ", "codubica, codpsn, codcpc ", "codsec, codpsn, codcpc ", "codpsn, codcpc ")
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "codcco, codpsn, orden", "codafp, codpsn, orden", "codubica, codpsn, orden", "codsec, codpsn, orden", "codpsn, orden")
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nSecuencia = 0
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    ' Genero los registros de la tabla de reporte
    ReDim a_Detalle(UBound(a_Columnas, 2))
    ReDim a_Niveles(nNivel)
    ReDim a_Quiebre(UBound(a_Detalle))
    ReDim a_Totales(UBound(a_Quiebre))
    sSeccion = Choose(nTabIndex + 1, "codcco", "codafp", "codubica", "codsec", "codrpt")
    a_Campos = Array("seccion", "secuencia", "cabecera1", "cabecera2", "detalle1", "detalle2")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    While Not porstRecordset.EOF
      sQuiebre = porstRecordset(sSeccion)
      ' Rango - quiebre adicional
      If nTabIndex <> 4 Then
        sDesquiebre = porstRecordset(Choose(nTabIndex + 1, "detcco", "desafp", "desubica", "dessec"))
        For n_Index = 1 To UBound(a_Quiebre): a_Quiebre(n_Index) = 0: Next n_Index
        ' Inserto el titulo del quiebre
        nSecuencia = nSecuencia + 1
        sDetalle = ""
        sDetalle = UCase(tdbSeleccion(nTabIndex).Caption) & ": " & Trim(sQuiebre) & " - " & Trim(sDesquiebre)
        a_Valores = Array("Q", nSecuencia, Left(sCabecera, 255), Mid(sCabecera, 256, 255), Left(sDetalle, 255), Mid(sDetalle, 256, 255))
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      End If
      Do
        ' Inicializo las variables
        sPersonal = porstRecordset("codpsn")
        sDetalle = ""
        For n_Index = 1 To UBound(a_Niveles): a_Niveles(n_Index) = 0: Next n_Index
        For n_Index = 1 To UBound(a_Detalle): a_Detalle(n_Index) = 0: Next n_Index
        Do
          a_Detalle(porstRecordset("orden")) = CDec(porstRecordset("nimporte"))
          porstRecordset.MoveNext
          If porstRecordset.EOF Then Exit Do
        Loop While sPersonal = porstRecordset("codpsn")
        porstRecordset.MovePrevious
        ' Genero los detalles por columna
        For n_Index = 1 To UBound(a_Columnas, 2)
          ' Registro de detalle
          If a_Columnas(3, n_Index) = "D" Then
            sValor = gdl_Funcion.PadR(Left(porstRecordset("dato_" & Format(n_Index, "00")), a_Columnas(7, n_Index)), a_Columnas(7, n_Index), Chr(32))
          ElseIf a_Columnas(3, n_Index) = "C" Then
            nValor = CDec(a_Detalle(n_Index))
            sValor = IIf(nValor = 0, " - ", Right(FormatNumber(nValor, 2), a_Columnas(7, n_Index)))
            sValor = gdl_Funcion.PadL(sValor, a_Columnas(7, n_Index), Chr(32))
            a_Totales(n_Index) = a_Totales(n_Index) + nValor
            a_Quiebre(n_Index) = a_Quiebre(n_Index) + nValor
            ' Sumo resto dependiende del signo
            If a_Columnas(5, n_Index) = "+" Then
              a_Niveles(a_Columnas(4, n_Index)) = a_Niveles(a_Columnas(4, n_Index)) + nValor
            Else
              a_Niveles(a_Columnas(4, n_Index)) = a_Niveles(a_Columnas(4, n_Index)) - nValor
            End If
          ElseIf a_Columnas(3, n_Index) = "A" Then
            nValor = a_Niveles(Val(a_Columnas(4, n_Index)) - 1)
            sValor = IIf(nValor = 0, " - ", Right(FormatNumber(nValor, 2), a_Columnas(7, n_Index)))
            sValor = gdl_Funcion.PadL(sValor, a_Columnas(7, n_Index), Chr(32))
            a_Totales(n_Index) = a_Totales(n_Index) + nValor
            a_Quiebre(n_Index) = a_Quiebre(n_Index) + nValor
            ' Sumo resto dependiende del signo
            If a_Columnas(5, n_Index) = "+" Then
              a_Niveles(a_Columnas(4, n_Index)) = a_Niveles(a_Columnas(4, n_Index)) + nValor
            Else
              a_Niveles(a_Columnas(4, n_Index)) = a_Niveles(a_Columnas(4, n_Index)) - nValor
            End If
          End If
          sDetalle = sDetalle & UCase(sValor)
        Next n_Index
        nSecuencia = nSecuencia + 1
        a_Valores = Array("D", nSecuencia, Left(sCabecera, 255), Mid(sCabecera, 256, 255), Left(sDetalle, 255), Mid(sDetalle, 256, 255))
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        porstRecordset.MoveNext
        ' Fin de archivo
        If porstRecordset.EOF Then Exit Do
      Loop While sQuiebre = porstRecordset(sSeccion)
      ' Totales rango - quiebre adicional
      If nTabIndex <> 4 Then
        ' Inserto los totales de centro de costo
        sDetalle = "x"
        For n_Index = 1 To UBound(a_Quiebre)
          ' Registro de detalle
          If a_Columnas(3, n_Index) = "D" Then
            sValor = gdl_Funcion.PadR("", a_Columnas(7, n_Index), Chr(32))
          Else
            nValor = CDec(a_Quiebre(n_Index))
            sValor = gdl_Funcion.PadL(Right(FormatNumber(nValor, 2), a_Columnas(7, n_Index)), a_Columnas(7, n_Index), Chr(32))
          End If
          sDetalle = sDetalle & UCase(sValor)
        Next n_Index
        nSecuencia = nSecuencia + 1
        a_Valores = Array("S", nSecuencia, Left(sCabecera, 255), Mid(sCabecera, 256, 255), Left(sDetalle, 255), Mid(sDetalle, 256, 255))
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      End If
    Wend
    ' Inserto los totales generales
    sDetalle = "x"
    For n_Index = 1 To UBound(a_Totales)
      ' Registro de detalle
      If a_Columnas(3, n_Index) = "D" Then
        sValor = gdl_Funcion.PadR("", a_Columnas(7, n_Index), Chr(32))
      Else
        nValor = CDec(a_Totales(n_Index))
        sValor = gdl_Funcion.PadL(Right(FormatNumber(nValor, 2), a_Columnas(7, n_Index)), a_Columnas(7, n_Index), Chr(32))
      End If
      sDetalle = sDetalle & UCase(sValor)
    Next n_Index
    a_Valores = Array("P", nSecuencia + 1, Left(sCabecera, 255), Mid(sCabecera, 256, 255), Left(sDetalle, 255), Mid(sDetalle, 256, 255))
    If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
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
Private Sub ReciboReporte(nTabIndex As Integer, s_Tabla As String, s_Reporte As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim a_Columnas(), a_Detalle(), sPersonal As String, nSecuencia As Long
  Dim a_Reporte(6), sMoneda As String, sMonedaPago As String
  Dim nImporteIng As Double, nImporteDsc As Double, nImportePag As Double
  Dim nIngreso As Integer, nDescuento As Integer
  Dim nContador As Integer, nColumna As Integer
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Imprimiendo Recibos ..."
  fMenu.panPercent.Visible = True
  sMoneda = IIf(s_Moneda = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt)

  ' Registros detalle con campos
  s_Sql = "SELECT rpt.alias, rpt.descripcion, rpt.orden, rpt.tipo, "
  s_Sql = s_Sql & "rpt.nivel, rpt.signo, rpt.impreso, rpt.longitud, vfx.valor "
  s_Sql = s_Sql & "FROM pldetareporte rpt "
  s_Sql = s_Sql & "LEFT JOIN plvarfunc vfx ON rpt.tipo=vfx.tipo AND rpt.alias=vfx.codigo AND IFNULL(vfx.valor, '')<>'' "
  s_Sql = s_Sql & "WHERE rpt.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND rpt.codrpt='" & s_Reporte & "' "
  s_Sql = s_Sql & "ORDER BY rpt.orden"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    n_Index = 0: s_Sql = ""
    While Not porstRecordset.EOF
      n_Index = n_Index + 1
      ' Redimensiono el arreglo de las cabeceras
      ReDim Preserve a_Columnas(8, n_Index)
      a_Columnas(1, n_Index) = Trim(porstRecordset("alias"))
      a_Columnas(7, n_Index) = CInt(porstRecordset("longitud"))
      a_Columnas(2, n_Index) = UCase(Left(Trim(porstRecordset("descripcion")), a_Columnas(7, n_Index)))
      a_Columnas(3, n_Index) = Trim(porstRecordset("tipo"))
      a_Columnas(4, n_Index) = Trim(porstRecordset("nivel"))
      a_Columnas(5, n_Index) = Trim(porstRecordset("signo"))
      a_Columnas(6, n_Index) = Trim(porstRecordset("impreso"))
      a_Columnas(8, n_Index) = gdl_Funcion.aTexto(porstRecordset("valor"))
      If a_Columnas(8, n_Index) <> "" And a_Columnas(8, n_Index) <> "NULL" Then
        s_Sql = s_Sql & Trim(porstRecordset!valor) & " AS dato_" & Format(n_Index, "00") & ", "
      End If
      porstRecordset.MoveNext
    Wend
  End If
  
  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT " & s_Sql
  s_Sql = s_Sql & "rpt.codrpt, rpt.alias, rpt.descripcion, rpt.orden, rpt.tipo, rpt.nivel, rpt.signo, rpt.impreso, rpt.longitud, "
  s_Sql = s_Sql & "res.codpsn, ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)), 2) AS nimporte, "
  s_Sql = s_Sql & "IFNULL(res.tipocpc, '2') AS tipocpc, dxr.codcco, cco.detcco, dxr.codafp, afp.desafp, dxr.codubica, ubi.desubica, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "me", "mn") & ", 0)), 2) AS importecmb, "
  s_Sql = s_Sql & "res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
  s_Sql = s_Sql & "IFNULL(dxr.codcgo, '') AS codcgo, IFNULL(cgo.descgo, '') AS descgo, dxr.fecingreso, IFNULL(psn.numdociden, '') AS numdociden, "
  s_Sql = s_Sql & "IFNULL(psn.numeroafp, '') AS numeroafp, IFNULL(dxr.codafp, '') AS codafp, IFNULL(afp.desafp, '') AS desafp, "
  s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', '" & s_Codmon_me & "', '" & s_Codmon_mn & "') AS monpago , IFNULL(pdo.despdo, '') AS despdo, pdo.tipocambio, "
  s_Sql = s_Sql & "dxr.codcco, cco.detcco, dxr.codafp, afp.desafp, dxr.codubica, ubi.desubica, dxr.codsec, sec.dessec "
  s_Sql = s_Sql & "FROM pldetareporte rpt "
  s_Sql = s_Sql & "INNER JOIN plresultado res ON rpt.codcls=res.codcls AND rpt.alias=res.codcpc AND rpt.tipo='C' "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON dxr.codeps=eps.codeps "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE rpt.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND rpt.codrpt='" & s_Reporte & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  If nTabIndex <> 4 Then
    s_Sql = s_Sql & "AND dxr." & Choose(nTabIndex + 1, "codcco", "codafp", "codubica", "codsec") & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "codcco, codpsn, codcpc ", "codafp, codpsn, codcpc ", "codubica, codpsn, codcpc ", "codsec, codpsn, codcpc ", "codpsn, codcpc ")
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "codcco, codpsn, orden", "codafp, codpsn, orden", "codubica, codpsn, orden", "codsec, codpsn, orden", "codpsn, orden")
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nSecuencia = 0
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    ' Genero los registros de la tabla de reporte
    a_Campos = Array("codpsn", "copia", "secuencia", "nompsn", "codcgo", "descgo", "fecingreso", "numdociden", "despdo", "numeroafp", "codafp", "desafp", "codcpcing", "descpcing", "impcpcing", "codcpcdsc", "descpcdsc", "impcpcdsc", "moneda", "monpago", "importipcmb", "impornetocmb")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero)
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    While Not porstRecordset.EOF
      sPersonal = porstRecordset("codpsn")
      nImporteIng = 0: nImporteDsc = 0: nImportePag = 0
      sMonedaPago = IIf(porstRecordset!monpago = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt)
      nContador = 0: nColumna = 6
      nIngreso = 0: nDescuento = 0
      ReDim a_Detalle(6, 0)
      Do
        nColumna = CInt(porstRecordset!tipocpc)
        If porstRecordset!Tipo = "C" Then
          nIngreso = nIngreso + IIf(nColumna = 0, 1, 0)
          nDescuento = nDescuento + IIf(nColumna = 1, 1, 0)
          nContador = Choose(nColumna + 1, nIngreso, nDescuento, 1)
          ' Redimensiono e inicializo el arreglo de los detalles
          If nContador > UBound(a_Detalle, 2) Then
            ReDim Preserve a_Detalle(6, nContador)
            a_Detalle(1, nContador) = "": a_Detalle(2, nContador) = ""
            a_Detalle(3, nContador) = "": a_Detalle(4, nContador) = ""
            a_Detalle(5, nContador) = CDec(0): a_Detalle(6, nContador) = CDec(0)
          End If
          ' Asigno los datos al arreglo
          a_Detalle(nColumna + 1, nContador) = porstRecordset("alias")
          a_Detalle(nColumna + 3, nContador) = porstRecordset("descripcion")
          a_Detalle(nColumna + 5, nContador) = CDec(porstRecordset!nImporte)
          ' Obtengo ingresos y descuentos otra moneda
          nImporteIng = nImporteIng + CDec(Choose(nColumna + 1, porstRecordset!importecmb, 0, 0))
          nImporteDsc = nImporteDsc + CDec(Choose(nColumna + 1, 0, porstRecordset!importecmb, 0))
        End If
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        DoEvents
        porstRecordset.MoveNext
        If porstRecordset.EOF Then Exit Do
      Loop While sPersonal = porstRecordset("codpsn")
      porstRecordset.MovePrevious
      ' Obtengo el importe en otra moneda
      nImportePag = CDec(nImporteIng - nImporteDsc)
      nImportePag = IIf(sMoneda = sMonedaPago, 0, nImportePag)
      ' Inserto los detalle maximo 10
      nContador = IIf(UBound(a_Detalle, 2) > 10, UBound(a_Detalle, 2), 10)
      For nSecuencia = 1 To nContador
        ' Inicializo los datos del detalle
        a_Reporte(1) = "": a_Reporte(2) = "": a_Reporte(5) = 0
        a_Reporte(3) = "": a_Reporte(4) = "": a_Reporte(6) = 0
        If UBound(a_Detalle, 2) >= nSecuencia Then
          a_Reporte(1) = a_Detalle(1, nSecuencia): a_Reporte(2) = a_Detalle(2, nSecuencia)
          a_Reporte(3) = a_Detalle(3, nSecuencia): a_Reporte(4) = a_Detalle(4, nSecuencia)
          a_Reporte(5) = a_Detalle(5, nSecuencia): a_Reporte(6) = a_Detalle(6, nSecuencia)
        End If
        a_Valores = Array(sPersonal, "0", nSecuencia, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), porstRecordset!numdociden, porstRecordset!despdo, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, a_Reporte(1), a_Reporte(3), a_Reporte(5), a_Reporte(2), a_Reporte(4), a_Reporte(6), sMonedaPago, sMonedaPago, CDec(porstRecordset!Tipocambio), nImportePag)
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      Next nSecuencia
      porstRecordset.MoveNext
    Wend
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
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
Private Sub RecuperaRegistros(ByVal nIndex As Integer, ByVal s_Orden As String)
  
  If nIndex = 0 Then
    s_Sql = "SELECT codcco, detcco, estcco "
    s_Sql = s_Sql & "FROM cocco "
    s_Sql = s_Sql & "WHERE LENGTH(codcco)=" & pn_NivelCenCosto & " "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 1 Then
    s_Sql = "SELECT codafp, desafp, estadoafp "
    s_Sql = s_Sql & "FROM plentidadafp "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 2 Then
    s_Sql = "SELECT codubica, desubica, estadoubica "
    s_Sql = s_Sql & "FROM plubicacion "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 3 Then
    s_Sql = "SELECT codsec, dessec, estadosec "
    s_Sql = s_Sql & "FROM plseccion "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 4 Then
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
  Dim nTabIndex As Integer
  Dim s_Periodo As String, s_Moneda As String
  Dim s_TituloReporte As String, s_FechaHora As String
  
  nTabIndex = tabRegister.Tab
  ' Verifico que Existan Registros
  If (dcaSeleccion(4).Recordset.EOF Or dcaSeleccion(4).Recordset.BOF) Or (dcaSeleccion(4).Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & tdbSeleccion(nTabIndex).Caption, vbExclamation: Exit Sub
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
    If txtFormato = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Formato de Reporte", vbExclamation: txtFormato.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Formato de Reporte no existe; verifique", vbExclamation: txtFormato.SetFocus: Exit Sub
    nTabIndex = IIf(chkPeriodos.Value, 4, tabRegister.Tab)
    ' Verifico que existan registros seleccionados
    If tdbSeleccion(4).SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(4).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 1 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 2 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 3 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    s_Periodo = ""
    nTabIndex = 4
    ' Barro el arreglo de registros (periodos) marcados (bookmarks)
    For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
      tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, "rptgralxpe", tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      s_Periodo = s_Periodo & " - " & Trim(tdbSeleccion(nTabIndex).Columns(1).Text)
    Next n_Index
    
    nTabIndex = IIf(chkPeriodos.Value, nTabIndex, tabRegister.Tab)
    If nTabIndex <> 4 Then
      ' Barro el arreglo de registros marcadas (bookmarks)
      For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
        tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
        gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, "rptgralxp" & nTabIndex, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      Next n_Index
    End If
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "REPORTE DE ANALISIS " & IIf(ribParametro(0).Value, "DETALLE", "RESUMEN")
    gdl_Procedure.ps_ReportName = IIf(ribFormato(0).Value, "rptpreplagnral" & LCase(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtFormato, "RH")), "rptrecipago")
    s_TituloReporte = gdl_Funcion.aTexto(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtFormato, "RT")) & IIf(ribFormato(0).Value, " - " & UCase(tdbSeleccion(nTabIndex).Caption), "")
    s_TituloReporte = s_TituloReporte & " (" & IIf(s_Moneda = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")"
    
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbSeleccion(nTabIndex).Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "":  aElemento(1, 2) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & s_TituloReporte & ";true"
    aElemento(2, 2) = "Periodo;" & s_Periodo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = ""
    aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If ribFormato(0).Value Then
      s_Sql = s_Sql & "secuencia smallint(5) Not Null, "
      s_Sql = s_Sql & "seccion char(1) Not Null, "
      s_Sql = s_Sql & "cabecera1 varchar(255) Null, "
      s_Sql = s_Sql & "cabecera2 varchar(255) Null, "
      s_Sql = s_Sql & "detalle1 varchar(255) Null, "
      s_Sql = s_Sql & "detalle2 varchar(255) Null, "
      s_Sql = s_Sql & "PRIMARY KEY (secuencia, seccion)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' Genera y selecciono la información del reporte
      GeneraReporte nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, txtFormato.Text, "rptgralxpe", s_FechaHora, s_Moneda
      s_Sql = "SELECT * "
      s_Sql = s_Sql & "FROM  tmp" & gdl_Procedure.ps_ReportName & " "
      s_Sql = s_Sql & "ORDER BY secuencia, seccion"
    Else
      aElemento(2, 2) = "Ruc;" & ps_RucEmpresa & ";true"
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, nompsn varchar(80) Null, "
      s_Sql = s_Sql & "codcgo char(3) Null, descgo varchar(80) Null, "
      s_Sql = s_Sql & "fecingreso date Null, numdociden varchar(11) Null, nroessalud varchar(15) Null, "
      s_Sql = s_Sql & "despdo varchar(40) Null, numeroafp varchar(15) Null, codafp char(2) Null, "
      s_Sql = s_Sql & "desafp varchar(40) Null, copia char(1) Not Null, secuencia smallint(3) Not Null, "
      s_Sql = s_Sql & "codcpcing varchar(4) Null, descpcing varchar(40) Null, impcpcing decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcdsc varchar(4) Null, descpcdsc varchar(40) Null, impcpcdsc decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "moneda char(3) Null, monpago char(3) Null, importipcmb decimal(6,3) Null Default '0', "
      s_Sql = s_Sql & "impornetocmb decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "PRIMARY KEY (codpsn, copia, secuencia))"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' Genera y selecciono la información del reporte
      ReciboReporte nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, txtFormato.Text, "rptgralxpe", s_FechaHora, s_Moneda
      s_Sql = "SELECT rec.*, cfg.regpatronal, "
      s_Sql = s_Sql & "CONCAT(IFNULL(cfg.repapepaterno, ''), ' ', IFNULL(cfg.repapematerno, ''), ', ', IFNULL(cfg.repnombres, '')) AS representante, "
      s_Sql = s_Sql & "cfg.logo, cfg.firma "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " rec, plcfgempresa cfg "
      s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "ORDER BY codpsn, copia, secuencia"
    End If
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, "rptgralxpe", "", ps_Usuario, s_FechaHora, "E"
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, "rptgralxp" & nTabIndex, "", ps_Usuario, s_FechaHora, "E"
    ' ]
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim s_TablaHelp As String
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Formatos reportes generados
    tdbHelp.Columns(0).DataField = "codrpt": tdbHelp.Columns(1).DataField = "desrpt"
    s_TablaHelp = "Formato Reporte"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("rpt", "codrpt", ps_ClsPlanilla, "")
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
  Me.Height = 6330: Me.Width = 11200
  Me.Left = 105: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  chkPeriodos.Value = True
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  
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
  For n_Index = 0 To 3
    aElemento(0, 1) = Choose(n_Index + 1, "codcco", "codafp", "codubica", "codsec")
    aElemento(1, 1) = Choose(n_Index + 1, "detcco", "desafp", "desubica", "dessec")
    aElemento(2, 1) = Choose(n_Index + 1, "estcco", "estadoafp", "estadoubica", "estadosec")
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
    gdl_Procedure.DefineStyleGrilla tdbSeleccion(n_Index), Choose(n_Index + 1, "Centro Costo", "Entidad Pensión", "Ubicación o Localidad", "Sección de Empresa"), 1
    ' Agrupacion de columnas y titulo DataView = dbgGroupView
    tdbSeleccion(n_Index).GroupByCaption = "Arrastrar titulo de columna de agrupación"
  Next n_Index
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(8, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ordenar Ascendente", "Ordenar Descendente", "Buscar Registro", "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&a", "&d", "&b", "&p", "&f", "&r", "&v", "&i")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 1
    ' Formato reporte
    ribFormato(n_Index).PictureUp = LoadPicture()
    ribFormato(n_Index).ToolTipText = "Formato " & Choose(n_Index + 1, "General", "Recibo")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "repogene", "reporeci") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribFormato(n_Index).PictureUp = LoadPicture(s_Sql)
    ' Tipo de analisis
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Analisis " & Choose(n_Index + 1, "Detallado", "Resumen")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "analmovs", "resumen") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribFormato(0).Value = True
  ribParametro(0).Value = True
 
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codbco", "desbco")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Entidad Bancaria", 2
  ']
 '[ Configuración de la grilla de registro
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
  gdl_Procedure.InicializaGrilla tdbSeleccion(4), aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbSeleccion(4).Columns(4).ValueItems.Presentation = dbgNormal
  tdbSeleccion(4).Columns(4).ValueItems.Translate = True
  For n_Index = 0 To 2
    tdbSeleccion(4).Columns(4).ValueItems.Add Item
    tdbSeleccion(4).Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Act, s_Estado_Ina, s_Estado_Blq)
    tdbSeleccion(4).Columns(4).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "proceok", "procenok", "perioblk") & ".bmp")
  Next n_Index
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbSeleccion(4), "Periodo Pago", 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbSeleccion(4).GroupByCaption = "Arrastrar titulo de columna de agrupación"
  ']
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  For n_Index = 0 To 4
    tdbSeleccion(n_Index).DataSource = dcaSeleccion(n_Index)
    RecuperaRegistros n_Index, tdbSeleccion(n_Index).Columns(0).DataField & " ASC"
  Next n_Index
  
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Formatos reportes generados
    txtFormato = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtFormato.SetFocus
  End Select
   
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Formatos reportes generados
    s_Sql = gdl_Funcion.HelpTablas("rpt", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
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
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaSeleccion(Index), tdbSeleccion(Index), " Registros"
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
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtFormato, "RP")
End Sub
