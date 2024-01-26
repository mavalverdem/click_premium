VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcFormatoBoleta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   11040
   Icon            =   "abcbolpago.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11040
   Begin TabDlg.SSTab tabRegister 
      Height          =   5100
      Left            =   75
      TabIndex        =   34
      Top             =   600
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   8996
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
      TabPicture(0)   =   "abcbolpago.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCuadro(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkCopia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmCuadro(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCodigo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDescripcion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbFonts"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtLinea"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbOrientacion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbCalidad"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.ComboBox cmbCalidad 
         Height          =   315
         ItemData        =   "abcbolpago.frx":0028
         Left            =   6525
         List            =   "abcbolpago.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1215
         Width           =   1300
      End
      Begin VB.ComboBox cmbOrientacion 
         Height          =   315
         ItemData        =   "abcbolpago.frx":002C
         Left            =   1340
         List            =   "abcbolpago.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1215
         Width           =   1300
      End
      Begin VB.TextBox txtLinea 
         Height          =   300
         Left            =   6870
         TabIndex        =   7
         Top             =   855
         Width           =   900
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   315
         ItemData        =   "abcbolpago.frx":0030
         Left            =   1340
         List            =   "abcbolpago.frx":0032
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   2475
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         TabIndex        =   3
         Top             =   480
         Width           =   5265
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1340
         MaxLength       =   8
         TabIndex        =   1
         Top             =   135
         Width           =   900
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3150
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   1530
         Width           =   9855
         _Version        =   65536
         _ExtentX        =   17383
         _ExtentY        =   5556
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
         Begin TrueOleDBGrid80.TDBDropDown tdbAyuda 
            Height          =   1500
            Left            =   1185
            TabIndex        =   19
            Top             =   1335
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   2646
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
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0)._GSX_SAVERECORDSELECTORS=   0
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1773"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1693"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   14215660
            ValueTranslate  =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
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
         Begin TrueOleDBGrid80.TDBGrid tdbDetalle 
            Height          =   3000
            Left            =   45
            TabIndex        =   18
            Top             =   105
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   5292
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=1958"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1879"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DataMode        =   4
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
            Height          =   3000
            Index           =   3
            Left            =   9210
            TabIndex        =   35
            Top             =   105
            Width           =   630
            _Version        =   65536
            _ExtentX        =   1111
            _ExtentY        =   5292
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
            Begin Threed.SSPanel panToolFmt 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   36
               Top             =   15
               Width           =   600
               _Version        =   65536
               _ExtentX        =   1058
               _ExtentY        =   450
               _StockProps     =   15
               Caption         =   "Detalle"
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
            Begin Threed.SSCommand cmdActionFmt 
               Height          =   360
               Index           =   0
               Left            =   105
               TabIndex        =   37
               Tag             =   "0"
               Top             =   525
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
               Picture         =   "abcbolpago.frx":0034
            End
            Begin Threed.SSCommand cmdActionFmt 
               Height          =   360
               Index           =   1
               Left            =   105
               TabIndex        =   38
               Tag             =   "0"
               Top             =   1290
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
               Picture         =   "abcbolpago.frx":0050
            End
            Begin Threed.SSCommand cmdActionFmt 
               Height          =   360
               Index           =   2
               Left            =   105
               TabIndex        =   39
               Tag             =   "0"
               Top             =   2025
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
               Picture         =   "abcbolpago.frx":006C
            End
         End
      End
      Begin Threed.SSCheck chkCopia 
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   825
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Linea Inicio Copia :"
         ForeColor       =   0
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
      Begin Threed.SSFrame frmCuadro 
         Height          =   1020
         Index           =   1
         Left            =   7845
         TabIndex        =   12
         Top             =   135
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   1799
         _StockProps     =   14
         Caption         =   " Medidas Papel  "
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
         Begin VB.TextBox txtPapel 
            Height          =   300
            Index           =   1
            Left            =   1140
            TabIndex        =   16
            Top             =   615
            Width           =   900
         End
         Begin VB.TextBox txtPapel 
            Height          =   300
            Index           =   0
            Left            =   1140
            TabIndex        =   14
            Top             =   270
            Width           =   900
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Alto :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   15
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Ancho :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   13
            Top             =   315
            Width           =   795
         End
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Calidad Impresión :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   5040
         TabIndex        =   10
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Orientación :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   8
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Font :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   870
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Código :"
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
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11040
      _Version        =   65536
      _ExtentX        =   19473
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
         Left            =   10020
         TabIndex        =   21
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
         Picture         =   "abcbolpago.frx":0088
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   9630
         TabIndex        =   22
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
         Picture         =   "abcbolpago.frx":00A4
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
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   8190
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   24
      Top             =   5775
      Width           =   11040
      _Version        =   65536
      _ExtentX        =   19473
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
         Left            =   6615
         TabIndex        =   25
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
         Picture         =   "abcbolpago.frx":00C0
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   6225
         TabIndex        =   26
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
         Picture         =   "abcbolpago.frx":00DC
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   4515
         TabIndex        =   27
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
         Picture         =   "abcbolpago.frx":00F8
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   4125
         TabIndex        =   28
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
         Picture         =   "abcbolpago.frx":0114
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   5100
      Index           =   0
      Left            =   10245
      TabIndex        =   29
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8996
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
         TabIndex        =   30
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
         TabIndex        =   31
         Tag             =   "0"
         Top             =   960
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
         Picture         =   "abcbolpago.frx":0130
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   32
         Tag             =   "0"
         Top             =   1770
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
         Picture         =   "abcbolpago.frx":014C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   33
         Tag             =   "0"
         Top             =   2550
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
         Picture         =   "abcbolpago.frx":0168
      End
   End
End
Attribute VB_Name = "fAbcFormatoBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private porstAyuda As ADODB.Recordset                   ' Recordset de ayuda
Private a_Formato As New XArrayDB                       ' Array de Formato de reporte
Private n_AnchoRpt As Integer, n_AltoRpt                ' Longitud, Altura del reporte acumulado
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  
  ' Detalle de Reporte
  cmdActionFmt(0).Enabled = (Me.Tag <> s_MdoData_Vis And tdbDetalle.Tag <> s_MdoData_Ins)
  cmdActionFmt(1).Enabled = (Me.Tag <> s_MdoData_Vis And (tdbDetalle.Tag = s_MdoData_Upd Or tdbDetalle.Tag = s_MdoData_Vis))
  cmdActionFmt(2).Enabled = (Me.Tag <> s_MdoData_Vis And (tdbDetalle.Tag = s_MdoData_Del Or tdbDetalle.Tag = s_MdoData_Vis))
  ' Tabla de formato detalle
  tdbDetalle.AllowAddNew = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  ' Bloqueo las columnas no editables
  For n_Index = 0 To 12
    tdbDetalle.Columns(n_Index).Locked = Not (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  Next n_Index
  tdbDetalle.Columns(3).Locked = True

End Sub
Private Function IndiceFuente(s_Fuente As String) As Integer
  Dim n_Contador As Integer, n_Fuente As Integer
  
  IndiceFuente = -1
  For n_Contador = 0 To cmbFonts.ListCount - 1
    ' Secciono el tipo de fuente
    If cmbFonts.List(n_Contador) = s_Fuente Then
      IndiceFuente = n_Contador
      Exit For
    End If
    ' Secciono el tipo de fuente por defecto
    If cmbFonts.List(n_Contador) = "Courier New" Then
      n_Fuente = n_Contador
    End If
  Next n_Contador
  ' Cuendo no se encontro el font
  If IndiceFuente = -1 Then IndiceFuente = n_Fuente
  
End Function
Private Sub RecuperaDetalle()
  
  ' Recupero el codigo del proceso
  n_AnchoRpt = 0
  ' Inicializo el arreglo
  a_Formato.ReDim 1, 0, 0, 17
  ' Genero la cadena de selección
  s_Sql = "SELECT bol.codboleta, bol.seccion, bol.origen, bol.dato, var.descripcion AS detalle, "
  s_Sql = s_Sql & "bol.tipodato, bol.fila, bol.columna, bol.longitud, bol.sizefont, bol.fontn, "
  s_Sql = s_Sql & "bol.fonts, bol.fontc, bol.desdato, bol.usrcre, bol.fyhcre, bol.usrmdf, bol.fyhmdf "
  s_Sql = s_Sql & "FROM pldetaboleta bol "
  s_Sql = s_Sql & "INNER JOIN plvarfunc var ON bol.dato=var.codigo AND bol.origen=var.tipo "
  s_Sql = s_Sql & "WHERE bol.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND bol.codboleta='" & Trim(txtCodigo.Text) & "' "
  s_Sql = s_Sql & "AND var.tipo='D' "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT bol.codboleta, bol.seccion, bol.origen, bol.dato, cpc.descpc AS detalle, "
  s_Sql = s_Sql & "bol.tipodato, bol.fila, bol.columna, bol.longitud, bol.sizefont, bol.fontn, "
  s_Sql = s_Sql & "bol.fonts, bol.fontc, bol.desdato, bol.usrcre, bol.fyhcre, bol.usrmdf, bol.fyhmdf "
  s_Sql = s_Sql & "FROM pldetaboleta bol "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON bol.dato=cpc.codcpc "
  s_Sql = s_Sql & "WHERE bol.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND bol.codboleta='" & Trim(txtCodigo.Text) & "' "
  s_Sql = s_Sql & "AND bol.origen='C' "
  s_Sql = s_Sql & "ORDER BY fila, columna"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    a_Formato.ReDim 1, porstRecordset.RecordCount, 0, 17
    n_Index = 0
    While Not porstRecordset.EOF
      n_Index = n_Index + 1
      a_Formato(n_Index, 0) = gdl_Funcion.aTexto(porstRecordset!seccion)
      a_Formato(n_Index, 1) = gdl_Funcion.aTexto(porstRecordset!origen)
      a_Formato(n_Index, 2) = gdl_Funcion.aTexto(porstRecordset!Dato)
      a_Formato(n_Index, 3) = gdl_Funcion.aTexto(porstRecordset!detalle)
      a_Formato(n_Index, 4) = gdl_Funcion.aTexto(porstRecordset!desdato)
      a_Formato(n_Index, 5) = gdl_Funcion.aTexto(porstRecordset("tipodato"))
      a_Formato(n_Index, 6) = CInt(porstRecordset!Fila)
      a_Formato(n_Index, 7) = CInt(porstRecordset!Columna)
      a_Formato(n_Index, 8) = CInt(porstRecordset!longitud)
      a_Formato(n_Index, 9) = Trim(porstRecordset!fontn)
      a_Formato(n_Index, 10) = Trim(porstRecordset!fontc)
      a_Formato(n_Index, 11) = Trim(porstRecordset("fonts"))
      a_Formato(n_Index, 12) = CDec(Trim(porstRecordset!sizefont))
      a_Formato(n_Index, 13) = gdl_Funcion.aTexto(porstRecordset!usrcre)
      a_Formato(n_Index, 14) = Format(porstRecordset!fyhcre, s_FmtFeHoMysql_0)
      a_Formato(n_Index, 15) = gdl_Funcion.aTexto(porstRecordset!usrmdf)
      a_Formato(n_Index, 16) = Format(porstRecordset!fyhmdf, s_FmtFeHoMysql_0)
      n_AnchoRpt = n_AnchoRpt + CInt(porstRecordset!longitud)
      porstRecordset.MoveNext
    Wend
  End If
  ' Cierro el recordset y saco del entorno
  porstRecordset.Close: Set porstRecordset = Nothing
  ' Asigno el arreglo a la grilla y relleno la misma
  Set tdbDetalle.Array = a_Formato
  tdbDetalle.ReBind
  
End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    n_Index = IndiceFuente("Courier New")
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fBoletaPago.dcaRegistro.Recordset!codboleta.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fBoletaPago.dcaRegistro.Recordset!desboleta.DefinedSize
    gdl_Procedure.EditCombo "AT", cmbFonts, n_Index, Me.Tag, False
    gdl_Procedure.EditCombo "AT", cmbOrientacion, 1, Me.Tag, False
    gdl_Procedure.EditCombo "AT", cmbCalidad, 1, Me.Tag, False
    gdl_Procedure.EditText "AT", txtPapel(0), FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtPapel(1), FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditOptionCheck "AT", chkCopia, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtLinea, 0, Me.Tag, False, 3, vbRightJustify
  Else
    n_Index = IndiceFuente(fBoletaPago.dcaRegistro.Recordset("font"))
    gdl_Procedure.EditText "PK", txtCodigo, fBoletaPago.dcaRegistro.Recordset!codboleta, Me.Tag, True, fBoletaPago.dcaRegistro.Recordset!codboleta.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fBoletaPago.dcaRegistro.Recordset!desboleta), Me.Tag, False, fBoletaPago.dcaRegistro.Recordset!desboleta.DefinedSize
    gdl_Procedure.EditCombo "AT", cmbFonts, n_Index, Me.Tag, False
    n_Index = CInt(fBoletaPago.dcaRegistro.Recordset!orientacion) - 1
    gdl_Procedure.EditCombo "AT", cmbOrientacion, n_Index, Me.Tag, False
    n_Index = Abs(fBoletaPago.dcaRegistro.Recordset!calidad) - 1
    gdl_Procedure.EditCombo "AT", cmbCalidad, n_Index, Me.Tag, False
    gdl_Procedure.EditText "AT", txtPapel(0), FormatNumber(fBoletaPago.dcaRegistro.Recordset!papelancho, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtPapel(1), FormatNumber(fBoletaPago.dcaRegistro.Recordset!papelalto, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtLinea, CInt(fBoletaPago.dcaRegistro.Recordset!lininicopia), Me.Tag, False, 3, vbRightJustify
    gdl_Procedure.EditOptionCheck "AT", chkCopia, (fBoletaPago.dcaRegistro.Recordset!copia = s_Estado_Act), Me.Tag, True
  End If
  ' Recupera información del detalle del reporte
  RecuperaDetalle

End Sub
Private Sub chkCopia_Click(Value As Integer)
  txtLinea.Locked = (Abs(Value) = vbUnchecked)
  txtLinea.Text = CInt(IIf(txtLinea.Locked, 0, txtLinea.Text))
End Sub
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  tdbDetalle.Tag = s_MdoData_Vis
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   txtDescripcion.SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim$(txtDescripcion) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim$(txtCodigo)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codboleta")
    a_Valores = Array(s_Registro)
    a_Tipos = Array(TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plboletapago", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fBoletaPago.dcaRegistro, fBoletaPago.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fBoletaPago.dcaRegistro.Recordset.EOF And fBoletaPago.dcaRegistro.Recordset.BOF) Or fBoletaPago.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fBoletaPago.dcaRegistro.Recordset.Find ("codboleta >= '" & s_Registro & "'")
      If fBoletaPago.dcaRegistro.Recordset.EOF Then fBoletaPago.dcaRegistro.Recordset.MoveLast
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
Private Sub cmdActionFmt_Click(Index As Integer)
  Dim oldBookmark As Variant

  ' Inicializo el modo de registro o selección
  Select Case Index
   Case 0     ' Añade un registro a la grilla
    n_Index = a_Formato.UpperBound(1) + 1
    a_Formato.ReDim 1, n_Index, 0, 16
    tdbDetalle.Bookmark = n_Index
    tdbDetalle.ReBind
   Case 1     ' Inserta un registro a la grilla
    If a_Formato.Count(1) = 0 And tdbDetalle.AllowAddNew Then
      a_Formato.ReDim 1, 1, 0, 16
      tdbDetalle.Refresh
      tdbDetalle.Bookmark = 1
    Else
      a_Formato.Insert 1, tdbDetalle.Bookmark
    End If
    tdbDetalle.ReBind
   Case 2     ' Elimina el registro actual de la grilla
    If a_Formato.Count(1) = 0 Then Beep: Exit Sub
    If tdbDetalle.Bookmark <> a_Formato.Count(1) Then
      a_Formato.Delete 1, tdbDetalle.Bookmark
      tdbDetalle.ReBind
    Else
      oldBookmark = tdbDetalle.Bookmark - 1
      a_Formato.Delete 1, tdbDetalle.Bookmark
      tdbDetalle.ReBind
      tdbDetalle.Bookmark = oldBookmark
    End If
  End Select
  tdbDetalle.SetFocus

End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero inicial, anterior, siguiente o final
  Select Case Index
   Case 0: fBoletaPago.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fBoletaPago.dcaRegistro.Recordset.BOF Then fBoletaPago.dcaRegistro.Recordset.MovePrevious
           If fBoletaPago.dcaRegistro.Recordset.BOF Then fBoletaPago.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fBoletaPago.dcaRegistro.Recordset.EOF Then fBoletaPago.dcaRegistro.Recordset.MoveNext
           If fBoletaPago.dcaRegistro.Recordset.EOF Then fBoletaPago.dcaRegistro.Recordset.MoveLast
   Case 3: fBoletaPago.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Copia As String * 1, s_Orientacion As String * 1
  Dim s_Seccion As String, s_Clase As String
  Dim s_Calidad As String
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If cmbFonts = "" Then Beep: MsgBox "Debe selecionar el tipo de fuente " & lblTitle, vbExclamation: cmbFonts.SetFocus: Exit Sub
  ' Dimensiones del papel
  If Not IsNumeric(txtPapel(0).Text) Then: MsgBox "Ancho ingresado no es correcto; Verifique", vbInformation: txtPapel(0).SetFocus: Exit Sub
  If CDec(txtPapel(0).Text) <= 0 Then Beep: MsgBox "Ancho no puede ser menor o igual a cero; Verifique", vbInformation: txtPapel(0).SetFocus: Exit Sub
  If Not IsNumeric(txtPapel(1).Text) Then: MsgBox "Alto ingresado no es correcto; Verifique", vbInformation: txtPapel(1).SetFocus: Exit Sub
  If CDec(txtPapel(1).Text) <= 0 Then Beep: MsgBox "Alto no puede ser menor o igual a cero; Verifique", vbInformation: txtPapel(1).SetFocus: Exit Sub
  ' Linea de inicio de copia
  If Not IsNumeric(txtLinea.Text) Then: MsgBox "Linea ingresada no es correcto; Verifique", vbInformation: txtLinea.SetFocus: Exit Sub
  If (CInt(txtLinea.Text) <= 0 And chkCopia.Value) Then Beep: MsgBox "Linia de incio de copia no puede ser meno o igual que cero; Verifique", vbInformation: txtLinea.SetFocus: Exit Sub

  ' Valido el detalle del formato
  For n_Index = a_Formato.LowerBound(1) To a_Formato.UpperBound(1)
    s_Seccion = IIf(a_Formato(n_Index, 0) = "C", "Cabecera", IIf(a_Formato(n_Index, 0) = "D", "Detalle", "Pie"))
    s_Clase = IIf(a_Formato(n_Index, 0) = "C", "Concepto", "Dato")
    If a_Formato(n_Index, 0) = "" Then
      Beep
      MsgBox "Debe Ingresare Sección de la Boleta, Fila: " & n_Index & ", Columna: 1", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If a_Formato(n_Index, 1) = "" Then
      Beep
      MsgBox "Debe Ingresar Clase de detalle de la Boleta, Fila: " & n_Index & ", Columna: 2", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If a_Formato(n_Index, 2) = "" Or a_Formato(n_Index, 3) = "" Then
      Beep
      MsgBox "Debe Ingresar Código de Detalle " & s_Seccion & "-" & s_Clase & ", Fila: " & n_Index & ", Columna: 3", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If a_Formato(n_Index, 5) = "" Then
      Beep
      MsgBox "Debe Ingresar Tipo de dato de " & s_Seccion & "-" & s_Clase & ", Fila: " & n_Index & ", Columna: 6", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If Not IsNumeric(a_Formato(n_Index, 6)) Or CInt(a_Formato(n_Index, 6)) <= 0 Then
      Beep
      MsgBox "Número de fila de dato de " & s_Seccion & "-" & s_Clase & " no es valida, Fila: " & n_Index & ", Columna: 7", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If Not IsNumeric(a_Formato(n_Index, 7)) Or CInt(a_Formato(n_Index, 7)) <= 0 Then
      Beep
      MsgBox "Número de columna de dato de " & s_Seccion & "-" & s_Clase & " no es valida, Fila: " & n_Index & ", Columna: 8", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If Not IsNumeric(a_Formato(n_Index, 8)) Or CInt(a_Formato(n_Index, 8)) <= 0 Then
      Beep
      MsgBox "Longirud de dato de " & s_Seccion & "-" & s_Clase & " no es valida, Fila: " & n_Index & ", Columna: 9", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If Not IsNumeric(a_Formato(n_Index, 12)) Or CDec(a_Formato(n_Index, 12)) <= 0 Then
      Beep
      MsgBox "Tamaño de font de dato de " & s_Seccion & "-" & s_Clase & " no es valida, Fila: " & n_Index & ", Columna: 13", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
  Next n_Index

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(txtCodigo)
  s_Copia = IIf(chkCopia.Value, s_Estado_Act, s_Estado_Ina)
  s_Orientacion = Trim(cmbOrientacion.ListIndex + 1)
  s_Calidad = Trim((cmbCalidad.ListIndex + 1) * -1)

  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  
  a_Campos = Array("codcls", "codboleta", "desboleta", "orientacion", "calidad", "papelancho", "papelalto", "font", "copia", "lininicopia", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, s_Registro, Trim(txtDescripcion), s_Orientacion, s_Calidad, CDec(txtPapel(0)), CDec(txtPapel(1)), Trim(cmbFonts), s_Copia, CInt(txtLinea), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codboleta")
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plboletapago", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plboletapago", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
    ' Elimino los registros del detalle del reporte
    If Not Records_Del("pldetaboleta", a_Where, a_Valores, a_Tipos) Then GoTo Error
  End If
  
  ' Realizo el proceso de actualización de los detalles
  For n_Index = a_Formato.LowerBound(1) To a_Formato.UpperBound(1)
    If Trim(a_Formato(n_Index, 0)) <> "" Then
      a_Campos = Array("codcls", "codboleta", "seccion", "origen", "dato", "desdato", "tipodato", "fila", "columna", "longitud", "fontn", "fontc", "fonts", "sizefont", "usrcre", "fyhcre", "usrmdf", "fyhmdf")
      a_Valores = Array(ps_ClsPlanilla, s_Registro, Trim(a_Formato(n_Index, 0)), Trim(a_Formato(n_Index, 1)), Trim(a_Formato(n_Index, 2)), Trim(a_Formato(n_Index, 4)), Trim(a_Formato(n_Index, 5)), CInt(a_Formato(n_Index, 6)), CInt(a_Formato(n_Index, 7)), CInt(a_Formato(n_Index, 8)), Trim(Abs(a_Formato(n_Index, 9))), Trim(Abs(a_Formato(n_Index, 10))), Trim(Abs(a_Formato(n_Index, 11))), CDec(a_Formato(n_Index, 12)), Trim(IIf(gdl_Funcion.aTexto(a_Formato(n_Index, 13)) = "", ps_Usuario, a_Formato(n_Index, 13))), Format(IIf(gdl_Funcion.aTexto(a_Formato(n_Index, 14)) = "", Now, a_Formato(n_Index, 14)), s_FmtFeHoMysql_0), Trim(IIf(Me.Tag = s_MdoData_Upd, ps_Usuario, "")), Format(IIf(Me.Tag = s_MdoData_Upd, Now, ""), s_FmtFeHoMysql_0))
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      If Not Records_Ins("pldetaboleta", a_Campos, a_Valores, a_Tipos) Then GoTo Error
    End If
  Next n_Index
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción

  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fBoletaPago.dcaRegistro, fBoletaPago.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fBoletaPago.dcaRegistro.Recordset.Find ("codboleta='" & s_Registro & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtCodigo.SetFocus
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
  Dim Item As New ValueItem
  Dim n_IndexValor As Integer

  'Establece Posición y Titulo del Formulario
  Me.Height = 6770: Me.Width = 11130
  Me.Left = 500: Me.Top = 100
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Boleta de Pago"
  lblTitle = "Formato de Boleta de Pago"
  ' Inicializo los datos de ayuda
  Set porstAyuda = New ADODB.Recordset
' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fBoletaPago.Tag
  tdbDetalle.Tag = s_MdoData_Vis

  ReDim aElemento(17, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Sección", "Clase", "Codigo", "Descripción", "Detalle", "Tipo", "Fila", "Columna", "Lng", "Negrita", "Cursiva", "Subrayada", "Tamaño", "usrcre", "fyhcre", "usrmdf", "fyhmdf")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "seccion", "origen", "dato", "descripcion", "desdato", "tipodato", "fila", "columna", "longitud", "fontn", "fontc", "fonts", "sizefont", "usrcre", "fyhcre", "usrmdf", "fyhmdf")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 780, 800, 800, 2220, 1500, 780, 500, 500, 500, 600, 600, 600, 600, 10, 10, 10, 10)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbCenter, vbCenter, vbCenter, vbRightJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "tdbAyuda", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1)
  Next n_Index
  
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbDetalle, aElemento, aElementos
  '[Cambio el formato de la grilla columna de valores
  ' Formato de tipo de sección
  tdbDetalle.Columns(0).ValueItems.Presentation = dbgComboBox
  tdbDetalle.Columns(0).ValueItems.Validate = True
  tdbDetalle.Columns(0).ValueItems.Translate = True
  ' Formato de tipo de Linea
  tdbDetalle.Columns(1).ValueItems.Presentation = dbgComboBox
  tdbDetalle.Columns(1).ValueItems.Validate = True
  tdbDetalle.Columns(1).ValueItems.Translate = True
  ' Formato de tipo de dato
  tdbDetalle.Columns(5).ValueItems.Presentation = dbgComboBox
  tdbDetalle.Columns(5).ValueItems.Validate = True
  tdbDetalle.Columns(5).ValueItems.Translate = True
  
  ' Formato de columna de font
  For n_Index = 9 To 11
    tdbDetalle.Columns(n_Index).ValueItems.Presentation = dbgCheckBox
    tdbDetalle.Columns(n_Index).ValueItems.Validate = True
  Next n_Index
  
  ' Combo de tipo de linea
  For n_Index = 0 To 1
   tdbDetalle.Columns(1).ValueItems.Add Item
   tdbDetalle.Columns(1).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "C", "D")
   tdbDetalle.Columns(1).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Concepto", "Dato")
  Next n_Index
  
  For n_Index = 0 To 2
    ' Combo de tipo de sección
    tdbDetalle.Columns(0).ValueItems.Add Item
    tdbDetalle.Columns(0).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "C", "D", "P")
    tdbDetalle.Columns(0).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Cabecera", "Detalle", "Pie")
    ' Combo de tipo de dato
    tdbDetalle.Columns(5).ValueItems.Add Item
    tdbDetalle.Columns(5).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "C", "N", "F")
    tdbDetalle.Columns(5).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Caracter", "Numero", "Fecha")
  Next n_Index
  ']
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbDetalle, "Diseño de Formato de Reporte", 3
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbDetalle.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Adiciono el listado de fonts, orientación, calidad
  For n_Index = 0 To Printer.FontCount - 1
    cmbFonts.AddItem Printer.Fonts(n_Index)
  Next n_Index
  For n_Index = 0 To 1
    cmbOrientacion.AddItem Choose(n_Index + 1, "Vertical", "Horizontal")
  Next n_Index
  For n_Index = 0 To 3
    cmbCalidad.AddItem Choose(n_Index + 1, "Borrador", "Baja", "Media", "Alta")
  Next n_Index
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 3)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&n", "&e", "&m")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionFmt, aElemento
  
  ' Cargo los graficos a los controles del toolbar de detalle
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anade", "inserta", "elimina")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Añadir ", "Inserta ", "Eliminar ") & "Detalle de Reporte"
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&n", "&i", "&e")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdActionFmt, aElemento

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
  l_ExistRecord = (fBoletaPago.dcaRegistro.Recordset.EOF Or fBoletaPago.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fBoletaPago.dcaRegistro.Recordset!codboleta

  ' Carga los datos en el formulario
  ShowScreen

 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcpc", "descpc")
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
  gdl_Procedure.InicializaGrilla tdbAyuda, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrillaDrop tdbAyuda
  ' Asigno el control de datos  ala grilla
  tdbAyuda.DataField = tdbAyuda.Columns(0).DataField
  tdbAyuda.ListField = tdbAyuda.Columns(0).DataField
  ' Recupero la información
  s_Sql = gdl_Funcion.HelpTablas("cpp", tdbAyuda.Columns(0).DataField, ps_ClsPlanilla & s_Registro, "")
  Set porstAyuda = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbAyuda.DataSource = porstAyuda
  ']

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub tdbDetalle_AfterColUpdate(ByVal ColIndex As Integer)
  ' Fuerzo a que se actualice la grilla y refresco
  tdbDetalle.Update
  tdbDetalle.Refresh
  tdbDetalle.SetFocus
End Sub
Private Sub tdbDetalle_AfterInsert()
  ' Elimino el registro aAñadido por la grilla
  a_Formato.Delete 1, a_Formato.UpperBound(1)
  ' Fuerzo a que se actualice la grilla y refresco
  tdbDetalle.ReBind
  tdbDetalle.Refresh
  tdbDetalle.SetFocus
End Sub
Private Sub tdbDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  
  ' Verifico que se encuentre en mantenimiento
  If Not tdbDetalle.AllowAddNew Then GoTo CancelaIngreso
  ' Verifico datos del formato
  If txtCodigo = "" Then MsgBox "Ingrese codigo de boleta", vbExclamation: txtCodigo.SetFocus: GoTo CancelaIngreso
  
  ' Verifico que primero se ingrese sección y tipo de item
  If ColIndex <> 0 Then
    If tdbDetalle.Columns(0).Text = "" Then
      Beep
      MsgBox "Primero ingreso tipo de sección", vbExclamation
      tdbDetalle.SetFocus
      GoTo CancelaIngreso
    End If
  End If
  If ColIndex > 1 Then
    If tdbDetalle.Columns(1).Text = "" Then
      Beep
      MsgBox "Primero ingreso tipo de detalle", vbExclamation
      tdbDetalle.SetFocus
      GoTo CancelaIngreso
    End If
  End If
  Exit Sub
  
CancelaIngreso:
  Cancel = True

End Sub
Private Sub tdbDetalle_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  Dim s_Descripcion As String
  Dim v_ValorActual As Variant
    
  If ColIndex <= 1 Then
    If tdbDetalle.Columns(ColIndex).Text = "" Then Beep: MsgBox "No se puede dejar en Blanco", vbExclamation: Cancel = True: Exit Sub
    If ColIndex = 0 Then tdbDetalle.Columns(1).Text = ""
    tdbDetalle.Columns(2).Text = ""
    tdbDetalle.Columns(3).Text = ""
    tdbDetalle.Columns(4).Text = ""
    tdbDetalle.Columns(5).Text = ""
    tdbDetalle.Columns(6).Text = "0"
    tdbDetalle.Columns(7).Text = "0"
    tdbDetalle.Columns(8).Text = "0"
    tdbDetalle.Columns(9).Text = "0"
    tdbDetalle.Columns(10).Text = "0"
    tdbDetalle.Columns(11).Text = "0"
    tdbDetalle.Columns(12).Text = "0"
  ElseIf ColIndex = 2 Then
    If tdbDetalle.Columns(ColIndex).Text = "" Then Beep: MsgBox "No se puede dejar en Blanco", vbExclamation: Cancel = True: Exit Sub
    If Left(tdbDetalle.Columns(1).Text, 1) = "C" Then
      s_Descripcion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, Trim(tdbDetalle.Columns(ColIndex).Text), "CP")
    ElseIf Left(tdbDetalle.Columns(1).Text, 1) = "D" Then
      s_Descripcion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, "D" & Trim(tdbDetalle.Columns(ColIndex).Text), "VC")
    End If
    If s_Descripcion = "???" Then
      Beep
      MsgBox tdbDetalle.Columns(1).Text & " ingresado No existe, por favor seleccione otra vez", vbExclamation
      Cancel = True
    Else
      tdbDetalle.Columns(3).Text = s_Descripcion
    End If
  ElseIf (ColIndex >= 6 And ColIndex <= 8) Or ColIndex = 12 Then
    If Not IsNumeric(tdbDetalle.Columns(ColIndex).Text) Then Beep: MsgBox "Debe Ingresar sólo valores númericos", vbExclamation: Cancel = True: Exit Sub
    If CInt(tdbDetalle.Columns(ColIndex).Text) < 0 Then Beep: MsgBox "Longitud debe ser mayor o igual a cero", vbExclamation: Cancel = True: Exit Sub
  End If

End Sub
Private Sub tdbDetalle_ButtonClick(ByVal ColIndex As Integer)
  Dim n_FilaActual As Variant
  
  ' Verifico datos de la boleta
  If txtCodigo = "" Or ColIndex <> 2 Then Exit Sub
  If (ColIndex = 2 And (tdbDetalle.Columns(0).Text = "" Or tdbDetalle.Columns(1).Text = "")) Then Exit Sub

  ' Obtengo la fila actual
  If gdl_Funcion.aTexto(tdbDetalle.Bookmark) = "" Then
    n_FilaActual = 1
  Else
    n_FilaActual = tdbDetalle.Bookmark
  End If
  
  ' Obtengo la sentencia de selección
  If tdbDetalle.Columns(1).Text = "Concepto" Then
    tdbAyuda.Columns(0).DataField = "codcpc": tdbAyuda.Columns(1).DataField = "descpc"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpt", "codcpc", ps_ClsPlanilla, "")
  ElseIf tdbDetalle.Columns(1).Text = "Dato" Then
    tdbAyuda.Columns(0).DataField = "codigo": tdbAyuda.Columns(1).DataField = "descripcion"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("vxf", "codigo", "TD", "")
  End If
  tdbAyuda.DataField = tdbAyuda.Columns(0).DataField
  tdbAyuda.ListField = tdbAyuda.Columns(0).DataField
  ' Recupera información
  Set porstAyuda = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbAyuda.DataSource = porstAyuda
  If porstAyuda.RecordCount = 0 Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  
End Sub
Private Sub tdbDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  ' Deshabilito la columna de descripción
  tdbDetalle.Columns(1).Button = Not (tdbDetalle.Columns(0).Text = "Acumulador")
  tdbDetalle.Columns(1).Locked = (tdbDetalle.Columns(0).Text = "Acumulador")
End Sub
Private Sub txtCodigo_GotFocus()
  gdl_Procedure.MarcaGet txtCodigo
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If txtCodigo = "" Then
      Beep
      MsgBox "Debe Ingresar el Código del " & lblTitle, vbExclamation
      txtCodigo.SetFocus
    Else
      txtDescripcion.SetFocus
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txtDescripcion_GotFocus()
  gdl_Procedure.MarcaGet txtDescripcion
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtLinea_GotFocus()
  gdl_Procedure.MarcaGet txtLinea
End Sub
Private Sub txtLinea_Validate(Cancel As Boolean)
  txtLinea.Text = IIf(Not IsNumeric(txtLinea.Text), 0, txtLinea.Text)
  txtLinea.Text = CInt(txtLinea.Text)
End Sub
Private Sub txtPapel_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPapel(Index)
End Sub
Private Sub txtPapel_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    If CDec(txtPapel(Index)) <= 0 Then
      Beep
      MsgBox "Debe Ingresar el Valor de la " & lblTitle, vbExclamation
      txtPapel(Index).SetFocus
    Else
      SendKeys "{TAB}"
    End If
    KeyAscii = 0
  End If
End Sub
Private Sub txtPapel_Validate(Index As Integer, Cancel As Boolean)
  txtPapel(Index).Text = IIf(Not IsNumeric(txtPapel(Index).Text), 0, txtPapel(Index).Text)
  If CInt(txtPapel(Index).Text) < 0 Then MsgBox "Dimensión de papel no puede ser negativo; Verifique", vbInformation: txtPapel(Index).SetFocus: Exit Sub
  txtPapel(Index).Text = FormatNumber(CDec(txtPapel(Index).Text), 2)
End Sub
