VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcPlanillaGnral 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5910
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   9105
   Icon            =   "abcpllagnral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   9105
   Begin TabDlg.SSTab tabRegister 
      Height          =   4740
      Left            =   75
      TabIndex        =   28
      Top             =   600
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   8361
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
      TabPicture(0)   =   "abcpllagnral.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkSeparador"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmCuadro(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDescripcion"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbPapel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbOrientacion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbSizeFont"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.ComboBox cmbSizeFont 
         Height          =   315
         ItemData        =   "abcpllagnral.frx":0028
         Left            =   7140
         List            =   "abcpllagnral.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   765
         Width           =   900
      End
      Begin VB.ComboBox cmbOrientacion 
         Height          =   315
         ItemData        =   "abcpllagnral.frx":002C
         Left            =   5310
         List            =   "abcpllagnral.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   765
         Width           =   1300
      End
      Begin VB.ComboBox cmbPapel 
         Height          =   315
         ItemData        =   "abcpllagnral.frx":0030
         Left            =   1340
         List            =   "abcpllagnral.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   765
         Width           =   2115
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   280
         Left            =   1340
         MaxLength       =   50
         TabIndex        =   3
         Top             =   435
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
         Height          =   280
         Left            =   1340
         MaxLength       =   8
         TabIndex        =   1
         Top             =   120
         Width           =   900
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3285
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1065
         Width           =   7920
         _Version        =   65536
         _ExtentX        =   13970
         _ExtentY        =   5794
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
         Begin Threed.SSPanel panToolBar 
            Height          =   3130
            Index           =   3
            Left            =   7275
            TabIndex        =   29
            Top             =   105
            Width           =   630
            _Version        =   65536
            _ExtentX        =   1111
            _ExtentY        =   5521
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
               TabIndex        =   30
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
               TabIndex        =   31
               Tag             =   "0"
               Top             =   750
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
               Picture         =   "abcpllagnral.frx":0034
            End
            Begin Threed.SSCommand cmdActionFmt 
               Height          =   360
               Index           =   1
               Left            =   105
               TabIndex        =   32
               Tag             =   "0"
               Top             =   1515
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
               Picture         =   "abcpllagnral.frx":0050
            End
            Begin Threed.SSCommand cmdActionFmt 
               Height          =   360
               Index           =   2
               Left            =   105
               TabIndex        =   33
               Tag             =   "0"
               Top             =   2250
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
               Picture         =   "abcpllagnral.frx":006C
            End
         End
         Begin TrueOleDBGrid80.TDBGrid tdbDetalle 
            Height          =   3130
            Left            =   45
            TabIndex        =   12
            Top             =   105
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   5530
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
            DataMode        =   4
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
         Begin TrueOleDBGrid80.TDBDropDown tdbAyuda 
            Height          =   1500
            Left            =   0
            TabIndex        =   13
            Top             =   45
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
      End
      Begin Threed.SSCheck chkSeparador 
         Height          =   285
         Left            =   5055
         TabIndex        =   4
         Top             =   120
         Width           =   1590
         _Version        =   65536
         _ExtentX        =   2805
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Imprime Cabecera"
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
      Begin VB.Label lblDato 
         Caption         =   "Tamaño Fuente :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   6810
         TabIndex        =   9
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblDato 
         Caption         =   "Orientación :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   4155
         TabIndex        =   7
         Top             =   795
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Papel :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   795
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
         Top             =   465
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
         Top             =   150
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9105
      _Version        =   65536
      _ExtentX        =   16060
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
         Left            =   8025
         TabIndex        =   15
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
         Picture         =   "abcpllagnral.frx":0088
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   7635
         TabIndex        =   16
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
         Picture         =   "abcpllagnral.frx":00A4
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
         TabIndex        =   17
         Top             =   120
         Width           =   6165
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   5400
      Width           =   9105
      _Version        =   65536
      _ExtentX        =   16060
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
         Left            =   5775
         TabIndex        =   19
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
         Picture         =   "abcpllagnral.frx":00C0
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   5385
         TabIndex        =   20
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
         Picture         =   "abcpllagnral.frx":00DC
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   3675
         TabIndex        =   21
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
         Picture         =   "abcpllagnral.frx":00F8
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   3285
         TabIndex        =   22
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
         Picture         =   "abcpllagnral.frx":0114
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4740
      Index           =   0
      Left            =   8310
      TabIndex        =   23
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8361
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         Picture         =   "abcpllagnral.frx":0130
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   26
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
         Picture         =   "abcpllagnral.frx":014C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   27
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
         Picture         =   "abcpllagnral.frx":0168
      End
   End
End
Attribute VB_Name = "fAbcPlanillaGnral"
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
Private n_Primary As Integer                            ' Clave de grilla
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
  For n_Index = 0 To 7
    tdbDetalle.Columns(n_Index).Locked = Not (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  Next n_Index
  tdbDetalle.Columns(1).Locked = True

End Sub
Private Sub RecuperaDetalle()
  
  ' Inicializo el arreglo
  a_Formato.ReDim 1, 0, 0, 13
  n_Primary = 0
  ' Genero la cadena de seleccion
  s_Sql = "SELECT codpll, fila, columna, despll, tipo, alias, descripcion, posicion, "
  s_Sql = s_Sql & "longitud, subrayado, usrcre, fyhcre, usrmdf, fyhmdf "
  s_Sql = s_Sql & "FROM plplanilla "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpll='" & Trim(txtCodigo.Text) & "' "
  s_Sql = s_Sql & "ORDER BY fila, columna, posicion"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    a_Formato.ReDim 1, porstRecordset.RecordCount, 0, 13
    n_Index = 0
    While Not porstRecordset.EOF
      n_Index = n_Index + 1
      n_Primary = CInt(porstRecordset!Columna)
      a_Formato(n_Index, 0) = CInt(porstRecordset!Fila)
      a_Formato(n_Index, 1) = CInt(porstRecordset!Columna)
      a_Formato(n_Index, 2) = CInt(porstRecordset!posicion)
      a_Formato(n_Index, 3) = gdl_Funcion.aTexto(porstRecordset!Tipo)
      a_Formato(n_Index, 4) = gdl_Funcion.aTexto(porstRecordset("alias"))
      a_Formato(n_Index, 5) = gdl_Funcion.aTexto(porstRecordset!descripcion)
      a_Formato(n_Index, 6) = gdl_Funcion.aTexto(porstRecordset!subrayado)
      a_Formato(n_Index, 7) = CInt(porstRecordset!longitud)
      a_Formato(n_Index, 8) = gdl_Funcion.aTexto(porstRecordset!usrcre)
      a_Formato(n_Index, 9) = Format(porstRecordset!fyhcre, s_FmtFeHoMysql_0)
      a_Formato(n_Index, 10) = gdl_Funcion.aTexto(porstRecordset!usrmdf)
      a_Formato(n_Index, 11) = Format(porstRecordset!fyhmdf, s_FmtFeHoMysql_0)
      a_Formato(n_Index, 12) = Format(porstRecordset!Fila, "00") & "-" & Format(n_Primary, "000")
      porstRecordset.MoveNext
    Wend
  End If
  ' Cierro el recordset y saco del entorno
  porstRecordset.Close: Set porstRecordset = Nothing
  ' Asigno el arreglo a la grilla y relleno la misma
  Set tdbDetalle.Array = a_Formato
  tdbDetalle.ReBind
  
  ' Actualizo tabla temporal detalle agrupación
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, "DROP TABLE IF EXISTS " & txtCodigo.Tag
  s_Sql = "CREATE TABLE IF NOT EXISTS " & txtCodigo.Tag & " "
  s_Sql = s_Sql & "SELECT det.codcls, det.codpll, det.fila, det.columna, det.codcpc, cpc.tipocpc, "
  s_Sql = s_Sql & "det.usrcre, det.fyhcre, det.usrmdf, det.fyhmdf "
  s_Sql = s_Sql & "FROM pldetaplanilla det "
  s_Sql = s_Sql & "INNER JOIN plconceplanilla cxp ON det.codcls=cxp.codcls AND det.codcpc=cxp.codcpc "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON det.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE det.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND det.codpll='" & Trim(txtCodigo.Text) & "' "
  s_Sql = s_Sql & "ORDER BY fila, columna, codcpc"
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  
End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fPlanillaGnral.dcaRegistro.Recordset!codpll.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fPlanillaGnral.dcaRegistro.Recordset!despll.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkSeparador, False, Me.Tag, True
    gdl_Procedure.EditCombo "AT", cmbPapel, 0, Me.Tag, False
    gdl_Procedure.EditCombo "AT", cmbOrientacion, 1, Me.Tag, False
    gdl_Procedure.EditCombo "AT", cmbSizeFont, 0, Me.Tag, False
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fPlanillaGnral.dcaRegistro.Recordset!codpll, Me.Tag, True, fPlanillaGnral.dcaRegistro.Recordset!codpll.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fPlanillaGnral.dcaRegistro.Recordset!despll), Me.Tag, False, fPlanillaGnral.dcaRegistro.Recordset!despll.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkSeparador, (fPlanillaGnral.dcaRegistro.Recordset!imprimecab = s_Estado_Act), Me.Tag, True
    n_Index = CInt(fPlanillaGnral.dcaRegistro.Recordset!sizepapel)
    gdl_Procedure.EditCombo "AT", cmbPapel, n_Index, Me.Tag, False
    n_Index = CInt(fPlanillaGnral.dcaRegistro.Recordset!posipapel) - 1
    gdl_Procedure.EditCombo "AT", cmbOrientacion, n_Index, Me.Tag, False
    n_Index = IIf(CInt(fPlanillaGnral.dcaRegistro.Recordset!sizefont) = 6, 0, IIf(CInt(fPlanillaGnral.dcaRegistro.Recordset!sizefont) = 8, 1, 2))
    gdl_Procedure.EditCombo "AT", cmbSizeFont, n_Index, Me.Tag, False
  End If
  ' Recupera información del detalle del reporte
  RecuperaDetalle

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
    a_Where = Array("codcls", "codpll")
    a_Valores = Array(ps_ClsPlanilla, s_Registro)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plgenreporte", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fPlanillaGnral.dcaRegistro, fPlanillaGnral.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fPlanillaGnral.dcaRegistro.Recordset.EOF And fPlanillaGnral.dcaRegistro.Recordset.BOF) Or fPlanillaGnral.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fPlanillaGnral.dcaRegistro.Recordset.Find ("codpll >= '" & s_Registro & "'")
      If fPlanillaGnral.dcaRegistro.Recordset.EOF Then fPlanillaGnral.dcaRegistro.Recordset.MoveLast
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
    a_Formato.ReDim 1, n_Index, 0, 13
    tdbDetalle.Bookmark = n_Index
    tdbDetalle.ReBind
   Case 1     ' Inserta un registro a la grilla
    If a_Formato.Count(1) = 0 And tdbDetalle.AllowAddNew Then
      a_Formato.ReDim 1, 1, 0, 13
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
   Case 0: fPlanillaGnral.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fPlanillaGnral.dcaRegistro.Recordset.BOF Then fPlanillaGnral.dcaRegistro.Recordset.MovePrevious
           If fPlanillaGnral.dcaRegistro.Recordset.BOF Then fPlanillaGnral.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fPlanillaGnral.dcaRegistro.Recordset.EOF Then fPlanillaGnral.dcaRegistro.Recordset.MoveNext
           If fPlanillaGnral.dcaRegistro.Recordset.EOF Then fPlanillaGnral.dcaRegistro.Recordset.MoveLast
   Case 3: fPlanillaGnral.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Tipo As String
  Dim s_ImpCabecera As String * 1, s_Subrayado As String * 1
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  ' Dimensiones del papel
  If cmbPapel = "" Then Beep: MsgBox "Debe selecionar el tipo de papel " & lblTitle, vbExclamation: cmbPapel.SetFocus: Exit Sub
  If cmbOrientacion = "" Then Beep: MsgBox "Debe selecionar la orientación de papel " & lblTitle, vbExclamation: cmbOrientacion.SetFocus: Exit Sub
  ' Tamaño de fuente
  If cmbSizeFont = "" Then Beep: MsgBox "Debe selecionar tamaño de carcater " & lblTitle, vbExclamation: cmbOrientacion.SetFocus: Exit Sub
  
  ' Valido el detalle del formato
  For n_Index = a_Formato.LowerBound(1) To a_Formato.UpperBound(1)
    s_Tipo = IIf(a_Formato(n_Index, 3) = "C", "Concepto", IIf(a_Formato(n_Index, 3) = "D", "Dato", "Grupo"))
    ' Verifico que primero se ingrese la fila
    If Not IsNumeric(a_Formato(n_Index, 0)) Or CInt(a_Formato(n_Index, 0)) <= 0 Then
      Beep
      MsgBox "Fila de " & s_Tipo & " no es valido, Fila: " & n_Index & ", Columna: 1", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    
    ' Verifico posicion
    If Not IsNumeric(a_Formato(n_Index, 2)) Or CInt(a_Formato(n_Index, 1)) < 0 Then
      Beep
      MsgBox "Posición de " & s_Tipo & " no es valido, Fila: " & n_Index & ", Columna: 2", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    
    ' Verifico que primero se ingrese el tipo de item
    If a_Formato(n_Index, 3) = "" Then
      Beep
      MsgBox "Debe Ingresar Tipo de Detalle, Fila: " & n_Index & ", Columna: 3", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
    If (a_Formato(n_Index, 3) = "C" Or a_Formato(n_Index, 3) = "D") And a_Formato(n_Index, 4) = "" Then
      Beep
      MsgBox "Debe Ingresar Código de " & s_Tipo & ", Fila: " & n_Index & ", Columna: 3", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
  
    ' Verifico si tiene detalle de agrupación
    If a_Formato(n_Index, 3) = "G" Then
      s_Sql = "SELECT IFNULL(COUNT(*), 0) AS registro "
      s_Sql = s_Sql & "FROM " & txtCodigo.Tag & " "
      s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND codpll='" & txtCodigo.Text & "' "
      s_Sql = s_Sql & "AND fila=" & CInt(Left(a_Formato(n_Index, 12), 2)) & " "
      s_Sql = s_Sql & "AND columna=" & CInt(Mid(a_Formato(n_Index, 12), 4))
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If porstRecordset!registro = 0 Then
        Beep
        MsgBox "Debe Ingresar Detalle Conceptos  " & s_Tipo & ", Fila: " & n_Index & ", Columna: 3", vbExclamation
        tdbDetalle.Bookmark = n_Index
        tdbDetalle.SetFocus
        Exit Sub
      End If
    End If
    
    ' Verifico la longitud
    If Not IsNumeric(a_Formato(n_Index, 7)) Or CInt(a_Formato(n_Index, 7)) <= 0 Then
      Beep
      MsgBox "Longitud de " & s_Tipo & " no es valido, Fila: " & n_Index & ", Columna: 7", vbExclamation
      tdbDetalle.Bookmark = n_Index
      tdbDetalle.SetFocus
      Exit Sub
    End If
  Next n_Index

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(txtCodigo)
  s_ImpCabecera = IIf(chkSeparador.Value, s_Estado_Act, s_Estado_Ina)
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  
  a_Campos = Array("codcls", "codpll", "despll", "sizefont", "sizepapel", "posipapel", "imprimecab", "fila", "columna", "posicion", "tipo", "alias", "descripcion", "longitud", "subrayado", "usrcre", "fyhcre", "usrmdf", "fyhmdf")
  a_Valores = Array(ps_ClsPlanilla, s_Registro, Trim(txtDescripcion), CDec(cmbSizeFont.Text), "", "", s_ImpCabecera, "fila", n_Index, "posicion", "tipo", "alias", "descripcion", "longitud", s_Subrayado, Trim(IIf(gdl_Funcion.aTexto(a_Formato(1, 8)) = "", ps_Usuario, a_Formato(1, 8))), Format(IIf(gdl_Funcion.aTexto(a_Formato(1, 9)) = "", Now, a_Formato(1, 9)), s_FmtFeHoMysql_0), Trim(IIf(Me.Tag = s_MdoData_Upd, ps_Usuario, "")), Format(IIf(Me.Tag = s_MdoData_Upd, Now, ""), s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpll")
  
  ' Elimino los registros del detalle del reporte
  If Not Records_Del("plplanilla", a_Where, a_Valores, a_Tipos) Then GoTo Error
  ' Realizo el proceso de actualización de los detalles
  For n_Index = a_Formato.LowerBound(1) To a_Formato.UpperBound(1)
    If Trim(a_Formato(n_Index, 0)) <> "" Then
      s_Subrayado = Trim(Abs(a_Formato(n_Index, 6)))
      a_Valores = Array(ps_ClsPlanilla, s_Registro, Trim(txtDescripcion), CDec(cmbSizeFont.Text), Trim(cmbPapel.ListIndex), Trim(cmbOrientacion.ListIndex + 1), s_ImpCabecera, CInt(a_Formato(n_Index, 0)), n_Index, CInt(a_Formato(n_Index, 2)), Trim(a_Formato(n_Index, 3)), Trim(a_Formato(n_Index, 4)), Trim(a_Formato(n_Index, 5)), CInt(a_Formato(n_Index, 7)), s_Subrayado, Trim(IIf(gdl_Funcion.aTexto(a_Formato(1, 8)) = "", ps_Usuario, a_Formato(1, 8))), Format(IIf(gdl_Funcion.aTexto(a_Formato(1, 9)) = "", Now, a_Formato(1, 9)), s_FmtFeHoMysql_0), Trim(IIf(Me.Tag = s_MdoData_Upd, ps_Usuario, "")), Format(IIf(Me.Tag = s_MdoData_Upd, Now, ""), s_FmtFeHoMysql_0))
      If Not Records_Ins("plplanilla", a_Campos, a_Valores, a_Tipos) Then GoTo Error
      If a_Formato(n_Index, 3) = "G" Then
        ' Realizo el proceso de actualización de las agrupaciones
        s_Sql = "INSERT INTO pldetaplanilla "
        s_Sql = s_Sql & "SELECT tmp.codcls, tmp.codpll, " & CInt(a_Formato(n_Index, 0)) & ", " & n_Index & ", tmp.codcpc, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM " & txtCodigo.Tag & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.codpll='" & txtCodigo.Text & "' "
        s_Sql = s_Sql & "AND tmp.fila=" & CInt(Left(a_Formato(n_Index, 12), 2)) & " "
        s_Sql = s_Sql & "AND tmp.columna=" & CInt(Mid(a_Formato(n_Index, 12), 4)) & " "
        s_Sql = s_Sql & "ORDER BY codcpc"
        If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Error
      End If
    End If
  Next n_Index
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción

  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fPlanillaGnral.dcaRegistro, fPlanillaGnral.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fPlanillaGnral.dcaRegistro.Recordset.Find ("codpll='" & s_Registro & "'")
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
  Me.Height = 6390: Me.Width = 9190
  Me.Left = 1700: Me.Top = 400
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Planilla"
  lblTitle = "Formato de Planilla"
  ' Inicializo los datos de ayuda
  Set porstAyuda = New ADODB.Recordset
' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fPlanillaGnral.Tag
  tdbDetalle.Tag = s_MdoData_Vis

  ReDim aElemento(13, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Fila", "Columna", "Posición", "Tipo", "Codigo", "Descripción", "Sub", "Lng", "usrcre", "fyhcre", "usrmdf", "fyhmdf", "clave")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "fila", "columna", "posicion", "tipo", "alias", "descripcion", "subrayado", "longitud", "usrcre", "fyhcre", "usrmdf", "fyhmdf", "clave")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 550, 10, 550, 1000, 900, 2810, 400, 400, 10, 10, 10, 10, 10)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbRightJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "", "", "", "", "", "", "", "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone, dbgMergeNone)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True, True, True, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "tdbAyuda", "", "", "", "", "", "", "", "")
    aElemento(n_Index, 9) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1)
  Next n_Index
  
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbDetalle, aElemento, aElementos
  '[Cambio el formato de la grilla columna de valores
  ' Formato de tipo de item
  tdbDetalle.Columns(3).ValueItems.Presentation = dbgComboBox
  tdbDetalle.Columns(3).ValueItems.Validate = True
  tdbDetalle.Columns(3).ValueItems.Translate = True
  ' Formato si item se subrayado o no
  tdbDetalle.Columns(6).ValueItems.Presentation = dbgCheckBox
  tdbDetalle.Columns(6).ValueItems.Validate = True
  For n_Index = 0 To 2
    tdbDetalle.Columns(3).ValueItems.Add Item
    tdbDetalle.Columns(3).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "C", "D", "G")
    tdbDetalle.Columns(3).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Concepto", "Dato", "Grupo")
  Next n_Index
  ']
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbDetalle, "Diseño de Formato de Planilla", 3
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbDetalle.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Adiciono el listado de papel, orientación, tamaño fonts
  For n_Index = 0 To 2
    cmbPapel.AddItem Choose(n_Index + 1, "A4 (210 x 297 mm)", "A3 (297 x 420 mm)", "Continuo USA estándar")
  Next n_Index
  For n_Index = 0 To 1
    cmbOrientacion.AddItem Choose(n_Index + 1, "Vertical", "Horizontal")
  Next n_Index
  For n_Index = 0 To 2
    cmbSizeFont.AddItem Choose(n_Index + 1, "6", "8", "10")
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
  l_ExistRecord = (fPlanillaGnral.dcaRegistro.Recordset.EOF Or fPlanillaGnral.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fPlanillaGnral.dcaRegistro.Recordset!codpll
  
  ' Obtengo el nombre de detalle de planilla
  txtCodigo.Tag = "grp" & Format(Now, "yyyymmddhhmmss")
  
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
Private Sub Form_Unload(Cancel As Integer)
  ' Elimino tabla temporal detalle agrupación
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, "DROP TABLE IF EXISTS " & txtCodigo.Tag
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
  ' Despues de actualizar columna
  Dim n_FilaActual As Variant
  
  ' Obtengo la fila actual
  If gdl_Funcion.aTexto(tdbDetalle.Bookmark) = "" Then
    n_FilaActual = 1
  Else
    n_FilaActual = tdbDetalle.Bookmark
  End If
  
  If ColIndex = 0 And n_FilaActual <> 1 Then
    If tdbDetalle.Columns(ColIndex).Text = a_Formato(n_FilaActual - 1, ColIndex) Then
      tdbDetalle.Columns(2).Text = CInt(a_Formato(n_FilaActual - 1, 2)) + CInt(a_Formato(n_FilaActual - 1, 7))
    End If
  End If
  
  ' Fuerzo a que se actualice la grilla y refresco
  tdbDetalle.Update
  tdbDetalle.Refresh
  tdbDetalle.SetFocus
End Sub
Private Sub tdbDetalle_AfterInsert()
  ' Despues de insertar registro
  ' Elimino el registro aAñadido por la grilla
  a_Formato.Delete 1, a_Formato.UpperBound(1)
  ' Fuerzo a que se actualice la grilla y refresco
  tdbDetalle.ReBind
  tdbDetalle.Refresh
  tdbDetalle.SetFocus
End Sub
Private Sub tdbDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  ' Antes de edición de columna
  
  ' Verifico que se encuentre en mantenimiento
  If Not tdbDetalle.AllowAddNew Then GoTo CancelaIngreso
  ' Verifico datos del reporte
  If txtCodigo = "" Then MsgBox "Ingrese codigo de Planilla", vbExclamation: txtCodigo.SetFocus: GoTo CancelaIngreso
  
  ' Verifico que primero se ingrese la fila
  If ColIndex <> 0 Then
    If tdbDetalle.Columns(0).Text = "" Then
      Beep
      MsgBox "Primero ingrese la fila de detalle", vbExclamation
      tdbDetalle.SetFocus
      GoTo CancelaIngreso
    End If
  End If
  ' Verifico que primero se ingrese la posición
  If ColIndex > 2 Then
    If tdbDetalle.Columns(2).Text = "" Then
      Beep
      MsgBox "Primero ingrese la posición incial de detalle", vbExclamation
      tdbDetalle.SetFocus
      GoTo CancelaIngreso
    End If
  End If
  ' Verifico que primero se ingrese el tipo de item
  If ColIndex > 3 Then
    If tdbDetalle.Columns(3).Text = "" Then
      Beep
      MsgBox "Primero ingrese tipo de detalle", vbExclamation
      tdbDetalle.SetFocus
      GoTo CancelaIngreso
    End If
  End If
  ' Verifico que primero se ingrese el codigo de item
  If ColIndex > 4 Then
    If tdbDetalle.Columns(4).Text = "" And tdbDetalle.Columns(3).Text <> "Grupo" Then
      Beep
      MsgBox "Primero ingrese codigo de detalle", vbExclamation
      tdbDetalle.SetFocus
      GoTo CancelaIngreso
    End If
  End If
  Exit Sub
  
CancelaIngreso:
  Cancel = True

End Sub
Private Sub tdbDetalle_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  ' Antes de actualización de columna
  Dim s_Descripcion As String
    
  If ColIndex = 0 Then
    If Not IsNumeric(tdbDetalle.Columns(ColIndex).Text) Then Beep: MsgBox "Debe Ingresar sólo valores númericos", vbExclamation: Cancel = True: Exit Sub
    If CInt(tdbDetalle.Columns(ColIndex).Text) <= 0 Then Beep: MsgBox "Fila debe ser mayor a cero", vbExclamation: Cancel = True: Exit Sub
    n_Primary = n_Primary + 1
    tdbDetalle.Columns(1).Text = "0"
    tdbDetalle.Columns(2).Text = "0"
    tdbDetalle.Columns(3).Text = ""
    tdbDetalle.Columns(4).Text = ""
    tdbDetalle.Columns(5).Text = ""
    tdbDetalle.Columns(6).Text = "0"
    tdbDetalle.Columns(7).Text = "0"
    tdbDetalle.Columns(12).Text = Format(tdbDetalle.Columns(ColIndex).Text, "00") & "-" & Format(n_Primary, "000")
  ElseIf ColIndex = 2 Then
    If Not IsNumeric(tdbDetalle.Columns(ColIndex).Text) Then Beep: MsgBox "Debe Ingresar sólo valores númericos", vbExclamation: Cancel = True: Exit Sub
    If CInt(tdbDetalle.Columns(ColIndex).Text) < 0 Then Beep: MsgBox "Posición debe ser mayor o igual a cero", vbExclamation: Cancel = True: Exit Sub
  ElseIf ColIndex = 3 Then
    If tdbDetalle.Columns(ColIndex).Text = "" Then Beep: MsgBox "No se puede dejar en Blanco", vbExclamation: Cancel = True: Exit Sub
    tdbDetalle.Columns(4).Text = ""
    tdbDetalle.Columns(5).Text = ""
    tdbDetalle.Columns(6).Text = "0"
    tdbDetalle.Columns(7).Text = "0"
  ElseIf ColIndex = 4 Then
    If (tdbDetalle.Columns(ColIndex).Text = "" And Left(tdbDetalle.Columns(0).Text, 1) <> "A") Then Beep: MsgBox "No se puede dejar en Blanco", vbExclamation: Cancel = True: Exit Sub
    If Left(tdbDetalle.Columns(3).Text, 1) = "C" Then
      s_Descripcion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, Trim(tdbDetalle.Columns(ColIndex).Text), "CP")
    ElseIf Left(tdbDetalle.Columns(3).Text, 1) = "D" Then
      s_Descripcion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, "D" & Trim(tdbDetalle.Columns(ColIndex).Text), "VC")
    End If
    If s_Descripcion = "???" Then
      Beep
      MsgBox tdbDetalle.Columns(3).Text & " ingresado No existe, por favor seleccione otra vez", vbExclamation
      Cancel = True
    Else
      tdbDetalle.Columns(5).Text = s_Descripcion
    End If
  ElseIf ColIndex = 7 Then
    If Not IsNumeric(tdbDetalle.Columns(ColIndex).Text) Then Beep: MsgBox "Debe Ingresar sólo valores númericos", vbExclamation: Cancel = True: Exit Sub
    If CInt(tdbDetalle.Columns(ColIndex).Text) <= 0 Then Beep: MsgBox "Longitud debe ser mayor a cero", vbExclamation: Cancel = True: Exit Sub
  End If

End Sub
Private Sub tdbDetalle_ButtonClick(ByVal ColIndex As Integer)
  Dim n_FilaActual As Variant
  
  ' Verifico si se encuentra mantenimiento y datos del reporte
  If Not tdbDetalle.AllowAddNew Then Exit Sub
  If txtCodigo = "" Or ColIndex <> 4 Then Exit Sub
  If (ColIndex = 4 And tdbDetalle.Columns(3).Text = "") Then Exit Sub

  ' Obtengo la fila actual
  If gdl_Funcion.aTexto(tdbDetalle.Bookmark) = "" Then
    n_FilaActual = 1
  Else
    n_FilaActual = tdbDetalle.Bookmark
  End If
  
  ' Detalle agrupador
  If tdbDetalle.Columns(3).Text = "Grupo" Then
    fPlanillaDeta.Show vbModal
    Exit Sub
  End If
  
  If tdbDetalle.Columns(3).Text = "Concepto" Then
    tdbAyuda.Columns(0).DataField = "codcpc": tdbAyuda.Columns(1).DataField = "descpc"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpt", "codcpc", ps_ClsPlanilla, "")
  ElseIf tdbDetalle.Columns(3).Text = "Dato" Then
    tdbAyuda.Columns(0).DataField = "codigo": tdbAyuda.Columns(1).DataField = "descripcion"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("vxf", "codigo", "ND", "")
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
  ' Deshabilito la columna de codigo si es detalle agrupador
  tdbDetalle.Columns(4).Locked = (tdbDetalle.Columns(3).Text = "Grupo")
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
