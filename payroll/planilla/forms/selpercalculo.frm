VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPeriCalculo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depuración de Cálculo"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "selpercalculo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7740
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   3675
      Left            =   45
      TabIndex        =   9
      Top             =   585
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6482
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2858"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2778"
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
   Begin MSAdodcLib.Adodc dcaRegistro 
      Height          =   330
      Left            =   45
      Top             =   4290
      Width           =   6840
      _ExtentX        =   12065
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
      Height          =   4050
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   585
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   7144
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
         TabIndex        =   5
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
         Picture         =   "selpercalculo.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1905
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
         Picture         =   "selpercalculo.frx":0028
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   8
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
         Top             =   720
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
         Picture         =   "selpercalculo.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   2325
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
         Picture         =   "selpercalculo.frx":0060
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
      Begin VB.TextBox txtProceso 
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
         TabIndex        =   10
         Top             =   105
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
         Left            =   2565
         TabIndex        =   11
         Top             =   150
         Width           =   195
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso Calculo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1320
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2505
      TabIndex        =   12
      Top             =   300
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
Attribute VB_Name = "fSelPeriCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_OptRegistro As String                         ' Instancia del formulario activo
'[
Private Sub DepuraPlanilla(ByVal s_Periodo As String)
  Dim s_OldMessage As String, nRegistros As Long
  Dim s_FechaHora As String
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Procesando Depuración de Cálculo ..."
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Obtengo la fecha de proceso
  s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
  ' Selecciono el personal a procesar marcados o general
  If s_OptRegistro = "inipllgral" Then
    s_Sql = "INSERT INTO rangoimpresion (proceso, valor, usrcre, fyhcre) "
    s_Sql = s_Sql & "SELECT DISTINCTROW '" & s_OptRegistro & "', "
    s_Sql = s_Sql & "res.codpsn, '" & ps_Usuario & "', '" & s_FechaHora & "' "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce='" & txtProceso.Text & "' "
    s_Sql = s_Sql & "AND res.codpdo='" & s_Periodo & "' "
    s_Sql = s_Sql & "ORDER BY res.codpsn"
    gdl_Conexion.Execucion s_Sql, Inserta
  ElseIf s_OptRegistro = "calpllpers" Then
    For nRegistros = 0 To o_CalculoPersona.tdbRegistro.SelBookmarks.Count - 1
      o_CalculoPersona.tdbRegistro.Bookmark = o_CalculoPersona.tdbRegistro.SelBookmarks(nRegistros)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, o_CalculoPersona.tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next nRegistros
  End If
  
  gdl_Conexion.IniciaTransaccion    'Inicia transacción
  
  ' Actualizo los periodos de cuenta corriente
  s_Sql = "UPDATE plcuentacte cte, plresultado res, rangoimpresion rng "
  s_Sql = s_Sql & "SET cte.abono_mn=IF(cte.codmon='" & s_Codmon_me & "', 0, cte.abono_mn), cte.abono_me=IF(cte.codmon='" & s_Codmon_mn & "', 0, cte.abono_me), "
  s_Sql = s_Sql & "cte.codpdocan=NULL, cte.estadoctacte='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE cte.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND cte.codpdoprv='" & s_Periodo & "' "
  s_Sql = s_Sql & "AND cte.numcuota<>0 "
  s_Sql = s_Sql & "AND res.codproce='" & txtProceso.Text & "' "
  s_Sql = s_Sql & "AND res.codcls=cte.codcls "
  s_Sql = s_Sql & "AND res.codpdo=cte.codpdoprv "
  s_Sql = s_Sql & "AND res.codpsn=cte.codpsn "
  s_Sql = s_Sql & "AND res.codcpc=cte.codcpc "
  s_Sql = s_Sql & "AND rng.valor=res.codpsn "
  s_Sql = s_Sql & "AND rng.proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion s_Sql, Modifica
  ' Incremento el porcentaje del proceso
  fMenu.panPercent.FloodPercent = 25
  
  ' Elimino los registros del resultado
  s_Sql = "DELETE res.* FROM plresultado res, rangoimpresion rng "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce='" & txtProceso.Text & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & s_Periodo & "' "
  s_Sql = s_Sql & "AND rng.valor=res.codpsn "
  s_Sql = s_Sql & "AND rng.proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND rng.fyhcre='" & s_FechaHora & "'"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
  ' Incremento el porcentaje del proceso
  fMenu.panPercent.FloodPercent = 50
  
  ' Verifico no existan datos procesados
  s_Sql = "SELECT COUNT(*) AS registros FROM plresultado "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo='" & s_Periodo & "' "
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nRegistros = CLng(porstRecordset!registros)
  porstRecordset.Close
  
  ' Elimino los datos de resultado
  s_Sql = "DELETE  dxr.* "
  s_Sql = s_Sql & "FROM pldatoresultado dxr "
  s_Sql = s_Sql & "WHERE dxr.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND dxr.codpdo='" & s_Periodo & "' "
  s_Sql = s_Sql & "AND NOT EXISTS(SELECT * FROM plresultado res "
  s_Sql = s_Sql & "WHERE res.codcls=dxr.codcls "
  s_Sql = s_Sql & "AND res.codpdo=dxr.codpdo "
  s_Sql = s_Sql & "AND res.codpsn=dxr.codpsn) "
  gdl_Conexion.Execucion s_Sql, Modifica
  ' Incremento el porcentaje del proceso
  fMenu.panPercent.FloodPercent = 75
  
  ' Actualizo el estado y fecha de proceso del periodo
  s_Sql = "UPDATE plperiodo "
  s_Sql = s_Sql & "SET estadopdo='" & IIf(nRegistros = 0, s_Estado_Ina, s_Estado_Act) & "' "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo='" & s_Periodo & "'"
  gdl_Conexion.Execucion s_Sql, Modifica
  ' Incremento el porcentaje del proceso
  fMenu.panPercent.FloodPercent = 100
  
  gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
  MsgBox "Se actualizo exitosamente " & s_TitleTable, vbInformation
  ' Refresco el Ado control y la grilla
  gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, s_TitleTable
  ' Verifico si aun existen registros
  dcaRegistro.Recordset.Find ("codpdo >= '" & s_Periodo & "'")
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  ' Elimino el rango de proceso
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codpdo, despdo, fechaini, fechafin, "
  s_Sql = s_Sql & "anopdo, mespdo, tpopdo, estadopdo "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly

End Sub
Private Sub cmdAction_Click(Index As Integer)
  
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0 ' Inicalización de Registros
    ' Verifico que Existan Registros
    If dcaRegistro.Recordset.RecordCount = 0 Then Beep: MsgBox "No Existen " & s_TitleTable & " para actualizar", vbExclamation: Exit Sub
    If tdbRegistro.SelBookmarks.Count > 1 Then Beep: MsgBox "Debe Seleccionar solo un periodo a depurar", vbExclamation: Exit Sub
    If gdl_Funcion.aTexto(dcaRegistro.Recordset!estadopdo) <> s_Estado_Act Then Beep: MsgBox s_TitleTable & " no se puede actualizar", vbExclamation: Exit Sub
    If txtProceso = "" Then Beep: MsgBox "Debe Ingresar el Procesos de calculo de planilla", vbExclamation: txtProceso.SetFocus: Exit Sub
    If (lblHelp(0) = "" Or lblHelp(0) = "???") Then Beep: MsgBox "Proceso de calculo de planilla no existe; verifique", vbExclamation: txtProceso.SetFocus: Exit Sub
    ' Verifico no existan periods de meses procesados posteriores
    s_Sql = "SELECT COUNT(*) AS registros FROM plperiodo pex "
    s_Sql = s_Sql & "WHERE pex.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND EXISTS(SELECT * FROM plperiodo pdo "
    s_Sql = s_Sql & "WHERE pdo.codcls=pex.codcls "
    s_Sql = s_Sql & "AND CONCAT(pex.anopdo, pex.mespdo)>CONCAT(pdo.anopdo, pdo.mespdo) "
    s_Sql = s_Sql & "AND pdo.codpdo='" & dcaRegistro.Recordset!codpdo & "') "
    s_Sql = s_Sql & "AND pex.estadopdo>='" & s_Estado_Act & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If porstRecordset!registros > 0 Then Beep: MsgBox "Actualizar " & s_TitleTable & " continuo", vbCritical: Exit Sub
    Beep
    If MsgBox("¿ Estás Seguro de Depurar / Inicializar el Periodo de Pago '" & Trim$(dcaRegistro.Recordset!despdo) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
      DepuraPlanilla dcaRegistro.Recordset!codpdo
    End If
   Case 1, 2  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 3 ' Busqueda de registro
    If Not (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Then
      Set go_tdbBusqueda = tdbRegistro
      Set go_dcaBusqueda = dcaRegistro
      gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
      fBusqueda.Show vbModal
    End If
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)

  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Periodo de pago bloqueados
    tdbHelp.Columns(0).DataField = "codproce": tdbHelp.Columns(1).DataField = "desproce"
    tdbHelp.Caption = "Proceso de Calculo"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("pro", "codproce", s_Estado_Act & ps_ClsPlanilla, "")
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
  Me.Height = 5160: Me.Width = 7830
  gdl_Procedure.CentraFormulario Me
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
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
    aElemento(n_Index, 2) = Choose(n_Index + 1, 900, 2980, 970, 970, 300)
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
  For n_Index = 0 To 1
    tdbRegistro.Columns(4).ValueItems.Add Item
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Act, s_Estado_Blq)
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "proceok", "perioblk") & ".bmp")
  Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(4, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "inicializa": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "borrafor", "ordascen", "orddesce", "busqueda")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Depuración de Calculo", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$)
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&c", "&a", "&d", "&b")
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
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  
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
    txtProceso = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtProceso.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Proceso de calculo
    s_Sql = gdl_Funcion.HelpTablas("pro", tdbHelp.Columns(ColIndex).DataField, s_Estado_Act & ps_ClsPlanilla, "")
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
Private Sub txtProceso_GotFocus()
  gdl_Procedure.MarcaGet txtProceso
End Sub
Private Sub txtProceso_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtProceso_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtProceso_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtProceso, "PC")
End Sub
