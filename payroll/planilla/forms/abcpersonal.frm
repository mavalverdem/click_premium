VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fAbcPersonal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9345
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   9225
   Icon            =   "abcpersonal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   9225
   Begin Threed.SSFrame frmRegister 
      Height          =   8175
      Left            =   0
      TabIndex        =   52
      Top             =   600
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   14420
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   13
         Left            =   -1.40000e5
         TabIndex        =   311
         Top             =   600
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbcomprobantes 
            Height          =   3465
            Left            =   135
            TabIndex        =   312
            Top             =   720
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   13
            Left            =   7200
            TabIndex        =   313
            Top             =   840
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel pantoolcomprobantes 
               Height          =   255
               Left            =   15
               TabIndex        =   314
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActioncomprobantes 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   315
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":000C
            End
            Begin Threed.SSCommand cmdActioncomprobantes 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   316
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":0028
            End
            Begin Threed.SSCommand cmdActioncomprobantes 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   317
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0044
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   318
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   12
         Left            =   -1.30000e5
         TabIndex        =   303
         Top             =   600
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbsuspension 
            Height          =   3465
            Left            =   135
            TabIndex        =   304
            Top             =   720
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   12
            Left            =   7200
            TabIndex        =   305
            Top             =   840
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolsuspension 
               Height          =   255
               Left            =   15
               TabIndex        =   306
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionsuspension 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   307
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":0060
            End
            Begin Threed.SSCommand cmdActionsuspension 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   308
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":007C
            End
            Begin Threed.SSCommand cmdActionsuspension 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   309
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0098
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   310
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   11
         Left            =   -1.20000e5
         TabIndex        =   295
         Top             =   480
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbterceros 
            Height          =   3465
            Left            =   135
            TabIndex        =   296
            Top             =   720
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   11
            Left            =   7200
            TabIndex        =   297
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolterceros 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   298
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionterceros 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   299
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":00B4
            End
            Begin Threed.SSCommand cmdActionterceros 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   300
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":00D0
            End
            Begin Threed.SSCommand cmdActionterceros 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   301
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":00EC
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   302
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   9
         Left            =   -1.00000e5
         TabIndex        =   287
         Top             =   600
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbempleador 
            Height          =   3465
            Left            =   135
            TabIndex        =   288
            Top             =   720
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   9
            Left            =   7155
            TabIndex        =   289
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolEmp 
               Height          =   255
               Left            =   15
               TabIndex        =   290
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionEmp 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   291
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":0108
            End
            Begin Threed.SSCommand cmdActionEmp 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   292
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":0124
            End
            Begin Threed.SSCommand cmdActionEmp 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   293
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0140
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   294
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   10
         Left            =   -11000
         TabIndex        =   279
         Top             =   600
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbesta 
            Height          =   3465
            Left            =   135
            TabIndex        =   280
            Top             =   720
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   10
            Left            =   7200
            TabIndex        =   281
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolEsta 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   282
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionEsta 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   283
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":015C
            End
            Begin Threed.SSCommand cmdActionEsta 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   284
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":0178
            End
            Begin Threed.SSCommand cmdActionEsta 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   285
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0194
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   286
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   0
         Left            =   -10000
         TabIndex        =   0
         Top             =   600
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame frmCuadro 
            Height          =   2985
            Index           =   0
            Left            =   75
            TabIndex        =   10
            Top             =   1140
            Width           =   5850
            _Version        =   65536
            _ExtentX        =   10319
            _ExtentY        =   5265
            _StockProps     =   14
            Caption         =   " Datos de Identificación"
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
            Begin VB.TextBox txtcorreo 
               Height          =   280
               Left            =   3000
               TabIndex        =   258
               Top             =   1680
               Width           =   2745
            End
            Begin VB.TextBox txtNacional 
               Height          =   280
               Left            =   1395
               TabIndex        =   28
               Top             =   2625
               Width           =   780
            End
            Begin VB.TextBox txtUbigeo 
               Height          =   280
               Index           =   0
               Left            =   180
               TabIndex        =   25
               Top             =   2280
               Width           =   975
            End
            Begin VB.TextBox txtNombres 
               Height          =   280
               Index           =   1
               Left            =   3015
               TabIndex        =   14
               Top             =   480
               Width           =   2700
            End
            Begin VB.TextBox txtNombres 
               Height          =   280
               Index           =   0
               Left            =   150
               TabIndex        =   12
               Top             =   480
               Width           =   2700
            End
            Begin VB.TextBox txtNombres 
               Height          =   280
               Index           =   2
               Left            =   150
               TabIndex        =   16
               Top             =   1065
               Width           =   2700
            End
            Begin VB.ComboBox cmbSexo 
               Height          =   315
               ItemData        =   "abcpersonal.frx":01B0
               Left            =   3015
               List            =   "abcpersonal.frx":01B2
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   1065
               Width           =   1260
            End
            Begin VB.ComboBox cmbEstadoCivil 
               Height          =   315
               ItemData        =   "abcpersonal.frx":01B4
               Left            =   4320
               List            =   "abcpersonal.frx":01B6
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1080
               Width           =   1485
            End
            Begin MSComCtl2.DTPicker dtpFecha 
               Height          =   285
               Left            =   150
               TabIndex        =   20
               Top             =   1665
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               _Version        =   393216
               Format          =   137560065
               CurrentDate     =   37515
            End
            Begin Threed.SSCommand cmdUbigeo 
               Height          =   300
               Index           =   0
               Left            =   1215
               TabIndex        =   55
               Top             =   2280
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   15
               Left            =   2280
               TabIndex        =   255
               Top             =   2640
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
            End
            Begin VB.Label lblDato 
               Caption         =   "Correo Electronico :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   49
               Left            =   3000
               TabIndex        =   257
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   15
               Left            =   2640
               TabIndex        =   256
               Top             =   2660
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Nacionalidad :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   47
               Left            =   180
               TabIndex        =   27
               Top             =   2670
               Width           =   1080
            End
            Begin VB.Label lblDato 
               Caption         =   "Lugar de Nacimiento :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   18
               Left            =   180
               TabIndex        =   24
               Top             =   2055
               Width           =   1755
            End
            Begin VB.Label lblUbigeo 
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
               Left            =   1620
               TabIndex        =   26
               Top             =   2340
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Nombres :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   150
               TabIndex        =   15
               Top             =   855
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Apellido Materno :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   4
               Left            =   3015
               TabIndex        =   13
               Top             =   255
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Apellido Paterno :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   150
               TabIndex        =   11
               Top             =   255
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Fecha de Nacimiento :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   150
               TabIndex        =   19
               Top             =   1455
               Width           =   1680
            End
            Begin VB.Label lblEdad 
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
               Left            =   1605
               TabIndex        =   21
               Top             =   1725
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Sexo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   3015
               TabIndex        =   17
               Top             =   855
               Width           =   1005
            End
            Begin VB.Label lblDato 
               Caption         =   "Estado Civil :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   4320
               TabIndex        =   22
               Top             =   840
               Width           =   1335
            End
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   1
            Left            =   6465
            TabIndex        =   41
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   0
            Left            =   4965
            TabIndex        =   39
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txtEssalud 
            Height          =   280
            Left            =   6030
            TabIndex        =   31
            Top             =   3690
            Width           =   1785
         End
         Begin VB.TextBox txtPorJudicial 
            Height          =   280
            Left            =   2925
            TabIndex        =   37
            Top             =   4440
            Width           =   800
         End
         Begin VB.TextBox txtDependientes 
            Height          =   280
            Left            =   1650
            TabIndex        =   35
            Top             =   4440
            Width           =   800
         End
         Begin VB.TextBox txtHijos 
            Height          =   280
            Left            =   120
            TabIndex        =   33
            Top             =   4440
            Width           =   800
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
            Left            =   120
            TabIndex        =   2
            Top             =   465
            Width           =   1320
         End
         Begin VB.TextBox txtTipoDocu 
            Height          =   280
            Left            =   1605
            TabIndex        =   4
            Top             =   450
            Width           =   500
         End
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Index           =   0
            Left            =   1605
            TabIndex        =   6
            Top             =   825
            Width           =   1440
         End
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Index           =   1
            Left            =   6030
            TabIndex        =   9
            Top             =   825
            Width           =   1530
         End
         Begin Threed.SSCheck chkExtanjero 
            Height          =   300
            Left            =   6030
            TabIndex        =   7
            Top             =   465
            Width           =   1650
            _Version        =   65536
            _ExtentX        =   2910
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "No Domiciliado"
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
         Begin Threed.SSCommand cmdHelp 
            Height          =   280
            Index           =   0
            Left            =   2160
            TabIndex        =   54
            Top             =   450
            Width           =   280
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCheck chkDsctoJudicial 
            Height          =   195
            Left            =   2925
            TabIndex        =   36
            Top             =   4200
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Descuento Judicial"
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
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Acreedor :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   6480
            TabIndex        =   40
            Top             =   4215
            Width           =   1290
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Deudor :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   45
            Left            =   4980
            TabIndex        =   38
            Top             =   4215
            Width           =   1170
         End
         Begin VB.Label lblDato 
            Caption         =   "Carnet Essalud :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   44
            Left            =   6030
            TabIndex        =   30
            Top             =   3450
            Width           =   1530
         End
         Begin VB.Label lblDato 
            Caption         =   "Dependientes :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   1650
            TabIndex        =   34
            Top             =   4215
            Width           =   1110
         End
         Begin VB.Label lblDato 
            Caption         =   "Número de Hijos :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   32
            Top             =   4215
            Width           =   1305
         End
         Begin VB.Label lblDato 
            Caption         =   "Fotografía :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   6030
            TabIndex        =   29
            Top             =   1170
            Width           =   1680
         End
         Begin VB.Label lblDato 
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
            Left            =   120
            TabIndex        =   1
            Top             =   210
            Width           =   960
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Left            =   2580
            TabIndex        =   5
            Top             =   495
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Documento de Identificación :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1605
            TabIndex        =   3
            Top             =   210
            Width           =   2070
         End
         Begin VB.Label lblDato 
            Caption         =   "Documento Militar :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   4530
            TabIndex        =   8
            Top             =   870
            Width           =   1380
         End
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Left            =   6030
            Stretch         =   -1  'True
            ToolTipText     =   "Haga doble click para fotografía"
            Top             =   1440
            Width           =   1800
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   8
         Left            =   -80000
         TabIndex        =   247
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbAnterior 
            Height          =   3465
            Left            =   135
            TabIndex        =   249
            Top             =   810
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   7
            Left            =   7155
            TabIndex        =   250
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolAnt 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   251
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionAnt 
               Height          =   360
               Index           =   0
               Left            =   150
               TabIndex        =   252
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":01B8
            End
            Begin Threed.SSCommand cmdActionAnt 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   253
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":01D4
            End
            Begin Threed.SSCommand cmdActionAnt 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   254
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":01F0
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   5
            Left            =   135
            TabIndex        =   248
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   7
         Left            =   -70000
         TabIndex        =   239
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbFamiliar 
            Height          =   3465
            Left            =   135
            TabIndex        =   241
            Top             =   810
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   6
            Left            =   7155
            TabIndex        =   242
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolFam 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   243
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionFam 
               Height          =   360
               Index           =   0
               Left            =   150
               TabIndex        =   244
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":020C
            End
            Begin Threed.SSCommand cmdActionFam 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   245
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":0228
            End
            Begin Threed.SSCommand cmdActionFam 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   246
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0244
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   4
            Left            =   135
            TabIndex        =   240
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   5
         Left            =   -50000
         TabIndex        =   231
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbEstudio 
            Height          =   3465
            Left            =   135
            TabIndex        =   233
            Top             =   810
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   4
            Left            =   7155
            TabIndex        =   234
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolEst 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   235
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionEst 
               Height          =   360
               Index           =   0
               Left            =   150
               TabIndex        =   236
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":0260
            End
            Begin Threed.SSCommand cmdActionEst 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   237
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":027C
            End
            Begin Threed.SSCommand cmdActionEst 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   238
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0298
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   232
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   4
         Left            =   -40000
         TabIndex        =   223
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbExperiencia 
            Height          =   3465
            Left            =   135
            TabIndex        =   225
            Top             =   810
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   3
            Left            =   7155
            TabIndex        =   226
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolExp 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   227
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionExp 
               Height          =   360
               Index           =   0
               Left            =   150
               TabIndex        =   228
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":02B4
            End
            Begin Threed.SSCommand cmdActionExp 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   229
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":02D0
            End
            Begin Threed.SSCommand cmdActionExp 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   230
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":02EC
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   224
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   6
         Left            =   -60000
         TabIndex        =   215
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbContrato 
            Height          =   3465
            Left            =   135
            TabIndex        =   217
            Top             =   810
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   6112
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            Height          =   3465
            Index           =   5
            Left            =   7155
            TabIndex        =   218
            Top             =   810
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   6112
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
            Begin Threed.SSPanel panToolCon 
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   219
               Top             =   15
               Width           =   720
               _Version        =   65536
               _ExtentX        =   1270
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
            Begin Threed.SSCommand cmdActionCon 
               Height          =   360
               Index           =   0
               Left            =   150
               TabIndex        =   220
               Tag             =   "0"
               Top             =   765
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
               Picture         =   "abcpersonal.frx":0308
            End
            Begin Threed.SSCommand cmdActionCon 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   221
               Tag             =   "0"
               Top             =   1530
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
               Picture         =   "abcpersonal.frx":0324
            End
            Begin Threed.SSCommand cmdActionCon 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   222
               Tag             =   "0"
               Top             =   2265
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
               Picture         =   "abcpersonal.frx":0340
            End
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   216
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   3
         Left            =   -30000
         TabIndex        =   194
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
            Height          =   2745
            Left            =   300
            TabIndex        =   196
            Top             =   525
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   4842
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
         Begin Threed.SSCheck chkRemuneNeta 
            Height          =   300
            Left            =   1950
            TabIndex        =   200
            Top             =   3300
            Width           =   225
            _Version        =   65536
            _ExtentX        =   397
            _ExtentY        =   529
            _StockProps     =   78
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
         Begin Threed.SSFrame frmCuadro 
            Height          =   1365
            Index           =   3
            Left            =   300
            TabIndex        =   199
            Top             =   3330
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   2408
            _StockProps     =   14
            Caption         =   " Remuneración Neta "
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
            Font3D          =   3
            ShadowStyle     =   1
            Begin VB.TextBox txtRemuNeta 
               Height          =   280
               Left            =   975
               TabIndex        =   204
               Top             =   645
               Width           =   1185
            End
            Begin VB.TextBox txtConcepto 
               Height          =   280
               Index           =   1
               Left            =   975
               TabIndex        =   206
               Top             =   990
               Width           =   555
            End
            Begin VB.TextBox txtConcepto 
               Height          =   280
               Index           =   0
               Left            =   975
               TabIndex        =   202
               Top             =   300
               Width           =   555
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   13
               Left            =   1605
               TabIndex        =   211
               Top             =   300
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   14
               Left            =   1605
               TabIndex        =   213
               Top             =   990
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
               Caption         =   "Importe :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   39
               Left            =   150
               TabIndex        =   203
               Top             =   645
               Width           =   735
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   14
               Left            =   1935
               TabIndex        =   214
               Top             =   1020
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Reajuste :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   35
               Left            =   150
               TabIndex        =   205
               Top             =   990
               Width           =   735
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   13
               Left            =   1935
               TabIndex        =   212
               Top             =   330
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Neto :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   33
               Left            =   150
               TabIndex        =   201
               Top             =   300
               Width           =   735
            End
         End
         Begin Threed.SSCheck chkRemIntegral 
            Height          =   195
            Index           =   2
            Left            =   4395
            TabIndex        =   210
            Top             =   4485
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Remuneración Integral  [C.T.S.]"
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
         Begin Threed.SSCheck chkRemIntegral 
            Height          =   195
            Index           =   1
            Left            =   4395
            TabIndex        =   209
            Top             =   4230
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Remuneración Integral  [Vacaciones]"
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
         Begin Threed.SSCheck chkRemIntegral 
            Height          =   195
            Index           =   0
            Left            =   4395
            TabIndex        =   208
            Top             =   3990
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Remuneración Integral  [Gratificación]"
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
         Begin Threed.SSCheck chkRemImprecisa 
            Height          =   195
            Left            =   4395
            TabIndex        =   207
            Top             =   3645
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Remuneración Principal Imprecisa"
            ForeColor       =   16711680
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
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   840
            Index           =   0
            Left            =   4260
            Shape           =   4  'Rounded Rectangle
            Top             =   3915
            Width           =   3435
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Remuneración :"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   28
            Left            =   4395
            TabIndex        =   197
            Top             =   3375
            Width           =   1710
         End
         Begin VB.Label lblTotalRemunera 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Left            =   6210
            TabIndex        =   198
            Top             =   3330
            Width           =   1245
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre :"
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
            Height          =   285
            Index           =   0
            Left            =   300
            TabIndex        =   195
            Top             =   180
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7020
         Index           =   1
         Left            =   -10000
         TabIndex        =   57
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   12382
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtTipoVia 
            Height          =   280
            Left            =   315
            TabIndex        =   59
            Top             =   630
            Width           =   500
         End
         Begin VB.TextBox txtNombreVia 
            Height          =   280
            Left            =   4260
            TabIndex        =   61
            Top             =   630
            Width           =   3470
         End
         Begin VB.TextBox txtNumero 
            Height          =   280
            Index           =   0
            Left            =   315
            TabIndex        =   63
            Top             =   1260
            Width           =   870
         End
         Begin VB.TextBox txtNumero 
            Height          =   280
            Index           =   1
            Left            =   1365
            TabIndex        =   65
            Top             =   1260
            Width           =   870
         End
         Begin VB.TextBox txtTipoZona 
            Height          =   280
            Left            =   315
            TabIndex        =   67
            Top             =   1875
            Width           =   500
         End
         Begin VB.TextBox txtNombreZona 
            Height          =   280
            Left            =   4260
            TabIndex        =   69
            Top             =   1875
            Width           =   3470
         End
         Begin VB.TextBox txtReferencia 
            Height          =   280
            Left            =   300
            MultiLine       =   -1  'True
            TabIndex        =   71
            Top             =   2490
            Width           =   7425
         End
         Begin VB.TextBox txtUbigeo 
            Height          =   280
            Index           =   1
            Left            =   300
            TabIndex        =   73
            Top             =   3120
            Width           =   975
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   1635
            Index           =   1
            Left            =   270
            TabIndex        =   75
            Top             =   3495
            Width           =   7410
            _Version        =   65536
            _ExtentX        =   13070
            _ExtentY        =   2884
            _StockProps     =   14
            Caption         =   " Números Teléfonicos "
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
            Begin VB.TextBox txtLargaDistancia 
               Height          =   280
               Left            =   180
               TabIndex        =   77
               Top             =   585
               Width           =   500
            End
            Begin VB.TextBox txtTelefono 
               Height          =   280
               Index           =   1
               Left            =   1800
               TabIndex        =   82
               Top             =   1230
               Width           =   1290
            End
            Begin VB.TextBox txtTelefono 
               Height          =   280
               Index           =   0
               Left            =   180
               TabIndex        =   80
               Top             =   1230
               Width           =   1290
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   26
               Left            =   715
               TabIndex        =   333
               Top             =   585
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
               Index           =   26
               Left            =   1120
               TabIndex        =   78
               Top             =   630
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Código LDN :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   63
               Left            =   180
               TabIndex        =   76
               Top             =   330
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Móvil :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   11
               Left            =   1800
               TabIndex        =   81
               Top             =   975
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Fijo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   16
               Left            =   180
               TabIndex        =   79
               Top             =   975
               Width           =   1335
            End
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   280
            Index           =   1
            Left            =   870
            TabIndex        =   83
            Top             =   630
            Width           =   280
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   280
            Index           =   2
            Left            =   870
            TabIndex        =   85
            Top             =   1875
            Width           =   280
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand cmdUbigeo 
            Height          =   285
            Index           =   1
            Left            =   1335
            TabIndex        =   87
            Top             =   3120
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Vía :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   315
            TabIndex        =   58
            Top             =   375
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombre de Vía :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   4260
            TabIndex        =   60
            Top             =   375
            Width           =   2280
         End
         Begin VB.Label lblDato 
            Caption         =   "Número :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   315
            TabIndex        =   62
            Top             =   1005
            Width           =   870
         End
         Begin VB.Label lblDato 
            Caption         =   "Interior :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   1365
            TabIndex        =   64
            Top             =   1005
            Width           =   870
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
            Left            =   1275
            TabIndex        =   84
            Top             =   675
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Zona :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   315
            TabIndex        =   66
            Top             =   1620
            Width           =   1335
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
            Index           =   2
            Left            =   1275
            TabIndex        =   86
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombre de Zona :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4260
            TabIndex        =   68
            Top             =   1620
            Width           =   2280
         End
         Begin VB.Label lblDato 
            Caption         =   "Referencia :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   300
            TabIndex        =   70
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Ubigeo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   300
            TabIndex        =   72
            Top             =   2865
            Width           =   1335
         End
         Begin VB.Label lblUbigeo 
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
            Left            =   1740
            TabIndex        =   74
            Top             =   3165
            Width           =   195
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   7635
         Index           =   2
         Left            =   -20000
         TabIndex        =   88
         Top             =   480
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   13467
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtJornadaLabor 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   6360
            TabIndex        =   94
            Top             =   390
            Width           =   660
         End
         Begin TabDlg.SSTab tabPago 
            Height          =   2520
            Left            =   105
            TabIndex        =   95
            Top             =   765
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   4445
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   5
            TabHeight       =   520
            ForeColor       =   16711680
            TabCaption(0)   =   "Modalidad"
            TabPicture(0)   =   "abcpersonal.frx":035C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblHelp(4)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblHelp(5)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblDato(22)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lblDato(36)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblDato(37)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lblHelp(3)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lblHelp(12)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lblDato(42)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lblDato(41)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "lblHelp(11)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "lblHelp(6)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "lblDato(23)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "lblDato(58)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "lblHelp(23)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "cmdHelp(23)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "cmdHelp(12)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "cmdHelp(11)"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "cmdHelp(6)"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "cmdHelp(4)"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "cmdHelp(5)"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "chkCargoConfi"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "cmdHelp(3)"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "cmbCargoConfi"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "txtCargo"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "txtProfesion"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "txtTipoTraba"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "txtSeccion"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "txtUbicacion"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "txtCenCosto"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "txtCondicion"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).ControlCount=   30
            TabCaption(1)   =   "Remuneración"
            TabPicture(1)   =   "abcpersonal.frx":0378
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblHelp(17)"
            Tab(1).Control(1)=   "lblDato(51)"
            Tab(1).Control(2)=   "lblDato(55)"
            Tab(1).Control(3)=   "lblHelp(20)"
            Tab(1).Control(4)=   "lblDato(24)"
            Tab(1).Control(5)=   "lblDato(27)"
            Tab(1).Control(6)=   "lblHelp(9)"
            Tab(1).Control(7)=   "lblDato(59)"
            Tab(1).Control(8)=   "lblHelp(24)"
            Tab(1).Control(9)=   "cmdHelp(24)"
            Tab(1).Control(10)=   "chkInterbank(0)"
            Tab(1).Control(11)=   "cmdHelp(17)"
            Tab(1).Control(12)=   "cmdHelp(20)"
            Tab(1).Control(13)=   "chkPagoDolar"
            Tab(1).Control(14)=   "cmdHelp(9)"
            Tab(1).Control(15)=   "txtTipPag"
            Tab(1).Control(16)=   "txtPeriodicidad"
            Tab(1).Control(17)=   "txtNroCuenta(0)"
            Tab(1).Control(18)=   "txtBanco(0)"
            Tab(1).Control(19)=   "txtBanco(1)"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "CTS"
            TabPicture(2)   =   "abcpersonal.frx":0394
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lblDato(26)"
            Tab(2).Control(1)=   "lblDato(25)"
            Tab(2).Control(2)=   "lblHelp(10)"
            Tab(2).Control(3)=   "lblHelp(25)"
            Tab(2).Control(4)=   "lblDato(60)"
            Tab(2).Control(5)=   "lblDato(62)"
            Tab(2).Control(6)=   "cmdHelp(25)"
            Tab(2).Control(7)=   "chkInterbank(1)"
            Tab(2).Control(8)=   "chkCtsDolar"
            Tab(2).Control(9)=   "chkCtsDeposito"
            Tab(2).Control(10)=   "cmdHelp(10)"
            Tab(2).Control(11)=   "txtNroCuenta(1)"
            Tab(2).Control(12)=   "txtBanco(2)"
            Tab(2).Control(13)=   "txtBanco(3)"
            Tab(2).Control(14)=   "txtNroCuenta(2)"
            Tab(2).ControlCount=   15
            TabCaption(3)   =   "Pensión"
            TabPicture(3)   =   "abcpersonal.frx":03B0
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lblDato(32)"
            Tab(3).Control(1)=   "lblDato(31)"
            Tab(3).Control(2)=   "lblHelp(8)"
            Tab(3).Control(3)=   "chkComisMixta"
            Tab(3).Control(4)=   "cmdHelp(8)"
            Tab(3).Control(5)=   "frmCuadro(5)"
            Tab(3).Control(6)=   "frmCuadro(4)"
            Tab(3).Control(7)=   "txtNumeroAfp"
            Tab(3).Control(8)=   "txtEntidadAfp"
            Tab(3).ControlCount=   9
            Begin VB.TextBox txtNroCuenta 
               Height          =   280
               Index           =   2
               Left            =   -69420
               MaxLength       =   20
               TabIndex        =   147
               Top             =   1695
               Width           =   1980
            End
            Begin VB.TextBox txtEntidadAfp 
               Height          =   280
               Left            =   -74835
               MaxLength       =   2
               TabIndex        =   149
               Top             =   600
               Width           =   500
            End
            Begin VB.TextBox txtNumeroAfp 
               Height          =   280
               Left            =   -69105
               TabIndex        =   152
               Top             =   600
               Width           =   1530
            End
            Begin VB.TextBox txtBanco 
               Height          =   280
               Index           =   1
               Left            =   -74790
               MaxLength       =   2
               TabIndex        =   133
               Top             =   2115
               Width           =   500
            End
            Begin VB.TextBox txtBanco 
               Height          =   280
               Index           =   3
               Left            =   -74790
               MaxLength       =   2
               TabIndex        =   144
               Top             =   1695
               Width           =   500
            End
            Begin VB.TextBox txtCondicion 
               Height          =   280
               Left            =   3645
               MaxLength       =   2
               TabIndex        =   114
               Top             =   1545
               Width           =   500
            End
            Begin VB.TextBox txtCenCosto 
               Height          =   280
               Left            =   3645
               MaxLength       =   20
               TabIndex        =   117
               Top             =   2100
               Width           =   900
            End
            Begin VB.TextBox txtUbicacion 
               Height          =   280
               Left            =   120
               MaxLength       =   2
               TabIndex        =   103
               Top             =   1650
               Width           =   500
            End
            Begin VB.TextBox txtSeccion 
               Height          =   280
               Left            =   120
               MaxLength       =   2
               TabIndex        =   106
               Top             =   2145
               Width           =   500
            End
            Begin VB.TextBox txtTipoTraba 
               Height          =   280
               Left            =   120
               MaxLength       =   2
               TabIndex        =   97
               Top             =   570
               Width           =   500
            End
            Begin VB.TextBox txtProfesion 
               Height          =   280
               Left            =   120
               MaxLength       =   6
               TabIndex        =   100
               Top             =   1110
               Width           =   735
            End
            Begin VB.TextBox txtCargo 
               Height          =   280
               Left            =   3645
               MaxLength       =   2
               TabIndex        =   111
               Top             =   960
               Width           =   500
            End
            Begin VB.ComboBox cmbCargoConfi 
               Height          =   315
               ItemData        =   "abcpersonal.frx":03CC
               Left            =   5580
               List            =   "abcpersonal.frx":03CE
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   435
               Width           =   2040
            End
            Begin VB.TextBox txtBanco 
               Height          =   280
               Index           =   0
               Left            =   -74790
               MaxLength       =   2
               TabIndex        =   127
               Top             =   1545
               Width           =   500
            End
            Begin VB.TextBox txtNroCuenta 
               Height          =   280
               Index           =   0
               Left            =   -69420
               MaxLength       =   20
               TabIndex        =   131
               Top             =   1545
               Width           =   1980
            End
            Begin VB.TextBox txtPeriodicidad 
               Height          =   280
               Left            =   -74790
               MaxLength       =   2
               TabIndex        =   121
               Top             =   975
               Width           =   375
            End
            Begin VB.TextBox txtTipPag 
               Height          =   280
               Left            =   -71280
               MaxLength       =   2
               TabIndex        =   124
               Top             =   975
               Width           =   375
            End
            Begin VB.TextBox txtBanco 
               Height          =   280
               Index           =   2
               Left            =   -74790
               MaxLength       =   2
               TabIndex        =   138
               Top             =   1050
               Width           =   500
            End
            Begin VB.TextBox txtNroCuenta 
               Height          =   280
               Index           =   1
               Left            =   -69420
               MaxLength       =   20
               TabIndex        =   142
               Top             =   1050
               Width           =   1980
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   10
               Left            =   -74250
               TabIndex        =   191
               Top             =   1050
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCheck chkCtsDeposito 
               Height          =   285
               Left            =   -72480
               TabIndex        =   136
               Top             =   420
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "Deposito de C.T.S."
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
            Begin Threed.SSCheck chkCtsDolar 
               Height          =   285
               Left            =   -74790
               TabIndex        =   135
               Top             =   420
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "C.T.S. Dólares"
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
            Begin Threed.SSCheck chkInterbank 
               Height          =   285
               Index           =   1
               Left            =   -69420
               TabIndex        =   140
               Top             =   420
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "Cuenta Interbancaria"
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   9
               Left            =   -74250
               TabIndex        =   189
               Top             =   1545
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCheck chkPagoDolar 
               Height          =   285
               Left            =   -74790
               TabIndex        =   119
               Top             =   420
               Width           =   2160
               _Version        =   65536
               _ExtentX        =   3810
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "Remuneración en Dólares"
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   20
               Left            =   -74370
               TabIndex        =   187
               Top             =   975
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   17
               Left            =   -70845
               TabIndex        =   188
               Top             =   975
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCheck chkInterbank 
               Height          =   285
               Index           =   0
               Left            =   -69420
               TabIndex        =   129
               Top             =   420
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "Interbancaria"
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   3
               Left            =   675
               TabIndex        =   180
               Top             =   570
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCheck chkCargoConfi 
               Height          =   285
               Left            =   3645
               TabIndex        =   108
               Top             =   420
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "Cargo de Confianza"
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   5
               Left            =   915
               TabIndex        =   181
               Top             =   1110
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   4
               Left            =   4200
               TabIndex        =   184
               Top             =   960
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   6
               Left            =   4620
               TabIndex        =   186
               Top             =   2100
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   11
               Left            =   675
               TabIndex        =   182
               Top             =   1650
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   12
               Left            =   675
               TabIndex        =   183
               Top             =   2145
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   23
               Left            =   4200
               TabIndex        =   185
               Top             =   1545
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   25
               Left            =   -74250
               TabIndex        =   192
               Top             =   1695
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   24
               Left            =   -74250
               TabIndex        =   190
               Top             =   2115
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSFrame frmCuadro 
               Height          =   945
               Index           =   4
               Left            =   -71220
               TabIndex        =   160
               Top             =   1455
               Width           =   3855
               _Version        =   65536
               _ExtentX        =   6800
               _ExtentY        =   1667
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
               Alignment       =   1
               Font3D          =   3
               ShadowStyle     =   1
               Begin VB.ComboBox cmbSCTRs 
                  Height          =   315
                  ItemData        =   "abcpersonal.frx":03D0
                  Left            =   165
                  List            =   "abcpersonal.frx":03D2
                  Style           =   2  'Dropdown List
                  TabIndex        =   165
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.ComboBox cmbSCTRp 
                  Height          =   315
                  ItemData        =   "abcpersonal.frx":03D4
                  Left            =   2295
                  List            =   "abcpersonal.frx":03D6
                  Style           =   2  'Dropdown List
                  TabIndex        =   167
                  Top             =   600
                  Width           =   1455
               End
               Begin Threed.SSCheck chk27252 
                  Height          =   225
                  Left            =   2550
                  TabIndex        =   163
                  Top             =   120
                  Width           =   1110
                  _Version        =   65536
                  _ExtentX        =   1958
                  _ExtentY        =   397
                  _StockProps     =   78
                  Caption         =   "Ley 27252"
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
               Begin Threed.SSCheck chkEssaludVida 
                  Height          =   225
                  Left            =   120
                  TabIndex        =   161
                  Top             =   105
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   397
                  _StockProps     =   78
                  Caption         =   "E. Vida"
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
               Begin Threed.SSCheck chkSindicato 
                  Height          =   225
                  Left            =   1200
                  TabIndex        =   162
                  Top             =   105
                  Width           =   1290
                  _Version        =   65536
                  _ExtentX        =   2275
                  _ExtentY        =   397
                  _StockProps     =   78
                  Caption         =   "Sindicalizado"
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
               Begin VB.Label Label1 
                  Caption         =   "SCTR Salud"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   255
                  Left            =   165
                  TabIndex        =   164
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label2 
                  Caption         =   "SCTR Pension"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   255
                  Left            =   2295
                  TabIndex        =   166
                  Top             =   360
                  Width           =   1455
               End
            End
            Begin Threed.SSFrame frmCuadro 
               Height          =   1425
               Index           =   5
               Left            =   -74895
               TabIndex        =   153
               Top             =   990
               Width           =   2775
               _Version        =   65536
               _ExtentX        =   4895
               _ExtentY        =   2514
               _StockProps     =   14
               Caption         =   " Régimen Pensionario "
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
               Begin Threed.SSOption optRegimen 
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   154
                  Top             =   285
                  Width           =   2445
                  _Version        =   65536
                  _ExtentX        =   4322
                  _ExtentY        =   353
                  _StockProps     =   78
                  Caption         =   "AFP - SPP"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption optRegimen 
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   155
                  Top             =   525
                  Width           =   2445
                  _Version        =   65536
                  _ExtentX        =   4322
                  _ExtentY        =   353
                  _StockProps     =   78
                  Caption         =   "ONP - SNP (D.L. 19990)"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption optRegimen 
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   156
                  Top             =   765
                  Width           =   2445
                  _Version        =   65536
                  _ExtentX        =   4322
                  _ExtentY        =   353
                  _StockProps     =   78
                  Caption         =   "D.L. 20530 y otros regimenes"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSMask.MaskEdBox mskFecInicio 
                  Height          =   280
                  Left            =   1335
                  TabIndex        =   158
                  Top             =   1020
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin VB.Label lblDato 
                  Caption         =   "Fecha Ingreso :"
                  ForeColor       =   &H00000000&
                  Height          =   190
                  Index           =   40
                  Left            =   120
                  TabIndex        =   157
                  Top             =   1050
                  Width           =   1125
               End
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   8
               Left            =   -74280
               TabIndex        =   332
               Top             =   600
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCheck chkComisMixta 
               Height          =   285
               Left            =   -69225
               TabIndex        =   159
               Top             =   1125
               Width           =   1755
               _Version        =   65536
               _ExtentX        =   3096
               _ExtentY        =   503
               _StockProps     =   78
               Caption         =   "AFP - Comisión Mixta"
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
            Begin VB.Label lblDato 
               Caption         =   "Cuenta Interbancaria :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   62
               Left            =   -69420
               TabIndex        =   146
               Top             =   1440
               Width           =   1965
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   8
               Left            =   -73950
               TabIndex        =   150
               Top             =   630
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Entidad de Pensión :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   31
               Left            =   -74835
               TabIndex        =   148
               Top             =   360
               Width           =   1545
            End
            Begin VB.Label lblDato 
               Caption         =   "Número de Pensión :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   32
               Left            =   -69105
               TabIndex        =   151
               Top             =   360
               Width           =   1530
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   24
               Left            =   -73905
               TabIndex        =   134
               Top             =   2115
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Entidad Inter - Bancaria :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   59
               Left            =   -74790
               TabIndex        =   132
               Top             =   1860
               Width           =   1845
            End
            Begin VB.Label lblDato 
               Caption         =   "Entidad Interbancaria :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   60
               Left            =   -74790
               TabIndex        =   143
               Top             =   1440
               Width           =   1845
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   25
               Left            =   -73905
               TabIndex        =   145
               Top             =   1695
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
               Index           =   23
               Left            =   4545
               TabIndex        =   115
               Top             =   1590
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Condición de Trabajo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   58
               Left            =   3645
               TabIndex        =   113
               Top             =   1290
               Width           =   1650
            End
            Begin VB.Label lblDato 
               Caption         =   "Centro de Costo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   23
               Left            =   3660
               TabIndex        =   116
               Top             =   1845
               Width           =   1470
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
               Index           =   6
               Left            =   4950
               TabIndex        =   118
               Top             =   2145
               Width           =   210
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
               Index           =   11
               Left            =   1005
               TabIndex        =   104
               Top             =   1650
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Ubicación :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   41
               Left            =   120
               TabIndex        =   102
               Top             =   1425
               Width           =   1470
            End
            Begin VB.Label lblDato 
               Caption         =   "Sección :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   42
               Left            =   120
               TabIndex        =   105
               Top             =   1935
               Width           =   1470
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
               Index           =   12
               Left            =   1005
               TabIndex        =   107
               Top             =   2190
               Width           =   195
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   3
               Left            =   1020
               TabIndex        =   98
               Top             =   615
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Tipo de Trabajador :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   37
               Left            =   120
               TabIndex        =   96
               Top             =   345
               Width           =   1470
            End
            Begin VB.Label lblDato 
               Caption         =   "Cargo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   36
               Left            =   3645
               TabIndex        =   110
               Top             =   705
               Width           =   1470
            End
            Begin VB.Label lblDato 
               Caption         =   "Profesión u Ocupación :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   99
               Top             =   885
               Width           =   1815
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
               Index           =   5
               Left            =   1275
               TabIndex        =   101
               Top             =   1155
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
               Index           =   4
               Left            =   4545
               TabIndex        =   112
               Top             =   1005
               Width           =   195
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   9
               Left            =   -73905
               TabIndex        =   128
               Top             =   1545
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Entidad Bancaria :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   27
               Left            =   -74790
               TabIndex        =   126
               Top             =   1320
               Width           =   1545
            End
            Begin VB.Label lblDato 
               Caption         =   "Número de Cuenta :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   24
               Left            =   -69420
               TabIndex        =   130
               Top             =   1320
               Width           =   1965
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   20
               Left            =   -74010
               TabIndex        =   122
               Top             =   975
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Periodicidad :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   55
               Left            =   -74790
               TabIndex        =   120
               Top             =   735
               Width           =   1065
            End
            Begin VB.Label lblDato 
               Alignment       =   2  'Center
               Caption         =   "Tipo de Pago :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   51
               Left            =   -71280
               TabIndex        =   123
               Top             =   735
               Width           =   1065
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   17
               Left            =   -70485
               TabIndex        =   125
               Top             =   975
               Width           =   195
            End
            Begin VB.Label lblHelp 
               AutoSize        =   -1  'True
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
               Index           =   10
               Left            =   -73905
               TabIndex        =   139
               Top             =   1050
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Entidad Bancaria :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   25
               Left            =   -74790
               TabIndex        =   137
               Top             =   795
               Width           =   1545
            End
            Begin VB.Label lblDato 
               Caption         =   "Número de Cuenta :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   26
               Left            =   -69420
               TabIndex        =   141
               Top             =   795
               Width           =   1965
            End
         End
         Begin VB.TextBox txttributacion 
            Alignment       =   2  'Center
            Height          =   280
            Left            =   5520
            MaxLength       =   2
            TabIndex        =   328
            Top             =   7230
            Width           =   500
         End
         Begin VB.TextBox txtcatocupacional 
            Alignment       =   2  'Center
            Height          =   280
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   326
            Top             =   7230
            Width           =   500
         End
         Begin VB.TextBox txtmodformativa 
            Height          =   280
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   274
            Top             =   6630
            Width           =   500
         End
         Begin VB.TextBox txtfinperiodo 
            Height          =   280
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   273
            Top             =   6105
            Width           =   500
         End
         Begin VB.ComboBox cmbforprofesional 
            Height          =   315
            ItemData        =   "abcpersonal.frx":03D8
            Left            =   240
            List            =   "abcpersonal.frx":03DA
            Style           =   2  'Dropdown List
            TabIndex        =   269
            Top             =   6630
            Width           =   2295
         End
         Begin VB.TextBox txtSiteps 
            Height          =   280
            Left            =   270
            MaxLength       =   20
            TabIndex        =   172
            Top             =   4275
            Width           =   375
         End
         Begin VB.TextBox txtCodEps 
            Height          =   280
            Left            =   270
            MaxLength       =   20
            TabIndex        =   169
            Top             =   3660
            Width           =   500
         End
         Begin VB.ComboBox cmbSituacion 
            Height          =   315
            ItemData        =   "abcpersonal.frx":03DC
            Left            =   6240
            List            =   "abcpersonal.frx":03DE
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   3645
            Width           =   1710
         End
         Begin MSComCtl2.DTPicker dtpFecIngreso 
            Height          =   285
            Left            =   120
            TabIndex        =   90
            Top             =   390
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Format          =   141950977
            CurrentDate     =   37515
         End
         Begin MSMask.MaskEdBox mskFecBaja 
            Height          =   285
            Left            =   6240
            TabIndex        =   179
            Top             =   4710
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   7
            Left            =   825
            TabIndex        =   193
            Top             =   3660
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin MSMask.MaskEdBox mskFecSitua 
            Height          =   285
            Left            =   6240
            TabIndex        =   177
            Top             =   4215
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin Threed.SSCheck chkRL 
            Height          =   285
            Left            =   240
            TabIndex        =   259
            Top             =   4710
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Regimen Laboral (Publico)"
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
         Begin Threed.SSCheck chkDIS 
            Height          =   285
            Left            =   240
            TabIndex        =   260
            Top             =   4950
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Discapacidad"
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
         Begin Threed.SSCheck chkREG 
            Height          =   285
            Left            =   240
            TabIndex        =   261
            Top             =   5430
            Width           =   7215
            _Version        =   65536
            _ExtentX        =   12726
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Sujeto a régimen alternativo, acumulativo o atípico de jornada de trabajo y descanso"
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
         Begin Threed.SSCheck chkMAX 
            Height          =   285
            Left            =   240
            TabIndex        =   263
            Top             =   5190
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Sujeto a jornada de trabajo máxima"
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
         Begin Threed.SSCheck chkNOC 
            Height          =   285
            Left            =   3120
            TabIndex        =   264
            Top             =   4710
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Trabajador sujeto a horario nocturno"
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
         Begin Threed.SSCheck chkOIQ 
            Height          =   285
            Left            =   3120
            TabIndex        =   265
            Top             =   5190
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Tiene otros ingresos de Quinta Categoría"
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
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   16
            Left            =   750
            TabIndex        =   266
            Top             =   4275
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCheck chkSM 
            Height          =   285
            Left            =   240
            TabIndex        =   267
            Top             =   5835
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Seguro Medico Essalud/Privado"
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
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   18
            Left            =   4560
            TabIndex        =   275
            Top             =   6105
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   19
            Left            =   4560
            TabIndex        =   276
            Top             =   6630
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCheck chkQUI 
            Height          =   285
            Left            =   3120
            TabIndex        =   262
            Top             =   4950
            Width           =   4680
            _Version        =   65536
            _ExtentX        =   8255
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Indicador de rentas de quinta categoría exoneradas-inafectas "
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
         Begin Threed.SSCheck chkMRF 
            Height          =   285
            Left            =   240
            TabIndex        =   268
            Top             =   6090
            Width           =   3060
            _Version        =   65536
            _ExtentX        =   5397
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Madre con Resp. Familiar"
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
         Begin Threed.SSCheck chkPE 
            Height          =   285
            Left            =   120
            TabIndex        =   324
            Top             =   7110
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "Afiliación Asegura tu Pensión"
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
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   21
            Left            =   3120
            TabIndex        =   327
            Top             =   7230
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   22
            Left            =   6120
            TabIndex        =   329
            Top             =   7230
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCheck chkReingreso 
            Height          =   285
            Left            =   4875
            TabIndex        =   92
            Top             =   390
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Reingreso"
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
         Begin VB.Label lblDato 
            Caption         =   "Jornada Laboral :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   6360
            TabIndex        =   93
            Top             =   150
            Width           =   1470
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Index           =   22
            Left            =   6480
            TabIndex        =   331
            Top             =   7230
            Width           =   195
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Index           =   21
            Left            =   3480
            TabIndex        =   330
            Top             =   7230
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Convenio Doble Tributación"
            Height          =   255
            Index           =   57
            Left            =   5520
            TabIndex        =   325
            Top             =   7035
            Width           =   2295
         End
         Begin VB.Label lblDato 
            Caption         =   "Categoria Ocupacional"
            Height          =   255
            Index           =   56
            Left            =   2520
            TabIndex        =   323
            Top             =   7035
            Width           =   2295
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   120
            X2              =   7800
            Y1              =   6990
            Y2              =   6990
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   3840
            X2              =   3840
            Y1              =   5790
            Y2              =   6990
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   5640
            X2              =   5520
            Y1              =   6150
            Y2              =   6030
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   5640
            X2              =   5520
            Y1              =   5910
            Y2              =   6030
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   5520
            X2              =   7920
            Y1              =   6030
            Y2              =   6030
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   7920
            X2              =   7920
            Y1              =   6020
            Y2              =   4920
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            X1              =   7560
            X2              =   7920
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Index           =   19
            Left            =   4920
            TabIndex        =   278
            Top             =   6630
            Width           =   195
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Index           =   18
            Left            =   4920
            TabIndex        =   277
            Top             =   6150
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Modalidad Formativa"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   54
            Left            =   3960
            TabIndex        =   272
            Top             =   6390
            Width           =   1695
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo Fin del Periodo"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   53
            Left            =   3960
            TabIndex        =   271
            Top             =   5865
            Width           =   1575
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Centro de Formación Profesional"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   52
            Left            =   240
            TabIndex        =   270
            Top             =   6390
            Width           =   3015
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   120
            X2              =   7800
            Y1              =   5790
            Y2              =   5790
         End
         Begin VB.Label lblDato 
            Caption         =   "Situacion Trabajador :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   270
            TabIndex        =   171
            Top             =   4050
            Width           =   1575
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Index           =   16
            Left            =   1110
            TabIndex        =   173
            Top             =   4275
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha :"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   48
            Left            =   6240
            TabIndex        =   176
            Top             =   3990
            Width           =   1335
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
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
            Index           =   7
            Left            =   1155
            TabIndex        =   170
            Top             =   3660
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Baja :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   43
            Left            =   6240
            TabIndex        =   178
            Top             =   4500
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "EPS :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   270
            TabIndex        =   168
            Top             =   3405
            Width           =   1470
         End
         Begin VB.Label lblDato 
            Caption         =   "Situación :"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   34
            Left            =   6240
            TabIndex        =   174
            Top             =   3405
            Width           =   1470
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Ingreso :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   120
            TabIndex        =   89
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lblTiempo 
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
            Left            =   1440
            TabIndex        =   91
            Top             =   450
            Width           =   195
         End
      End
      Begin MSComctlLib.TabStrip tasRegister 
         Height          =   5175
         Left            =   120
         TabIndex        =   53
         Top             =   165
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   9128
         Style           =   2
         Separators      =   -1  'True
         TabMinWidth     =   1176
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   14
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Datos Personales"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Datos de Domicilio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Datos de la Empresa"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Remuneración"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Experiencia Laboral"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Estudios Realizados"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Contratos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Datos Familiares"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ing/Dscto Anteriores"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Otros Empleadores"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Establecimientos donde Labora el Trabajador"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Personal de Terceros"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Datos de Suspensión de Cuarta Categoría"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalle de Comprobantes de Cuarta Categoría"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   9225
      _Version        =   65536
      _ExtentX        =   16272
      _ExtentY        =   1085
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
         Left            =   8415
         TabIndex        =   43
         Top             =   120
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
         Picture         =   "abcpersonal.frx":03E0
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   8025
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcpersonal.frx":03FC
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8640
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
               Picture         =   "abcpersonal.frx":0418
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   3
         Left            =   7410
         TabIndex        =   319
         Top             =   120
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
         Picture         =   "abcpersonal.frx":0572
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   7020
         TabIndex        =   320
         Top             =   120
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
         Picture         =   "abcpersonal.frx":058E
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   6360
         TabIndex        =   321
         Top             =   120
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
         Picture         =   "abcpersonal.frx":05AA
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   6000
         TabIndex        =   322
         Top             =   120
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
         Picture         =   "abcpersonal.frx":05C6
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   5760
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   46
      Top             =   8835
      Width           =   9225
      _Version        =   65536
      _ExtentX        =   16272
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
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   8115
      Index           =   0
      Left            =   8430
      TabIndex        =   47
      Top             =   660
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   14314
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
         TabIndex        =   48
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
         TabIndex        =   49
         Tag             =   "0"
         Top             =   1110
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
         Picture         =   "abcpersonal.frx":05E2
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   50
         Tag             =   "0"
         Top             =   1965
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
         Picture         =   "abcpersonal.frx":05FE
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   51
         Tag             =   "0"
         Top             =   2790
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
         Picture         =   "abcpersonal.frx":061A
      End
   End
   Begin MSAdodcLib.Adodc dcaHelp 
      Height          =   330
      Left            =   285
      Top             =   495
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2025
      TabIndex        =   56
      Top             =   480
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
Attribute VB_Name = "fAbcPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla
Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Cuonter As Integer, s_ParCodigo As String     ' Indice para bucle, y parametro de codigo
Private s_Personal As String                            ' Codigo del registro
Private n_IndexTabs As Integer                          ' indicce la pestaña del tab control
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private a_RemuneraDefa As New XArrayDB                  ' Array de remuneraciones por default
Private n_ImporteCol As Double                          ' Valor anterior de la columna
Dim cnn As ADODB.Connection
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Cuonter = 0 To 3: cmdMove(n_Cuonter).Visible = (Me.Tag = s_MdoData_Vis): Next n_Cuonter
  
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  
  For n_Cuonter = 0 To 26: cmdHelp(n_Cuonter).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd): Next n_Cuonter
  For n_Cuonter = 0 To 1: cmdUbigeo(n_Cuonter).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd): Next n_Cuonter
  
  tdbRegistro.Columns(3).Locked = Not (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  tdbRegistro.Columns(4).Locked = Not (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  
  ' Cuarta pestaña
  cmdActionExp(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbExperiencia.Tag <> s_MdoData_Ins)
  cmdActionExp(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbExperiencia.Tag = s_MdoData_Upd Or tdbExperiencia.Tag = s_MdoData_Vis))
  cmdActionExp(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbExperiencia.Tag = s_MdoData_Del Or tdbExperiencia.Tag = s_MdoData_Vis))
  ' Quinta pestaña
  cmdActionEst(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbEstudio.Tag <> s_MdoData_Ins)
  cmdActionEst(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbEstudio.Tag = s_MdoData_Upd Or tdbEstudio.Tag = s_MdoData_Vis))
  cmdActionEst(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbEstudio.Tag = s_MdoData_Del Or tdbEstudio.Tag = s_MdoData_Vis))
  ' Sexta pestaña
  cmdActionCon(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbContrato.Tag <> s_MdoData_Ins)
  cmdActionCon(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbContrato.Tag = s_MdoData_Upd Or tdbContrato.Tag = s_MdoData_Vis))
  cmdActionCon(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbContrato.Tag = s_MdoData_Del Or tdbContrato.Tag = s_MdoData_Vis))
  ' Septima pestaña
  cmdActionFam(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbFamiliar.Tag <> s_MdoData_Ins)
  cmdActionFam(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbFamiliar.Tag = s_MdoData_Upd Or tdbFamiliar.Tag = s_MdoData_Vis))
  cmdActionFam(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbFamiliar.Tag = s_MdoData_Del Or tdbFamiliar.Tag = s_MdoData_Vis))
  ' Octava pestaña
  cmdActionAnt(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbAnterior.Tag <> s_MdoData_Ins)
  cmdActionAnt(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbAnterior.Tag = s_MdoData_Upd Or tdbAnterior.Tag = s_MdoData_Vis))
  cmdActionAnt(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbAnterior.Tag = s_MdoData_Del Or tdbAnterior.Tag = s_MdoData_Vis))
  ' Novena pestaña
  cmdActionEmp(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbempleador.Tag <> s_MdoData_Ins)
  cmdActionEmp(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbempleador.Tag = s_MdoData_Upd Or tdbempleador.Tag = s_MdoData_Vis))
  cmdActionEmp(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbempleador.Tag = s_MdoData_Del Or tdbempleador.Tag = s_MdoData_Vis))
  ' Decima pestaña
  cmdActionEsta(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbesta.Tag <> s_MdoData_Ins)
  cmdActionEsta(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbesta.Tag = s_MdoData_Upd Or tdbesta.Tag = s_MdoData_Vis))
  cmdActionEsta(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbesta.Tag = s_MdoData_Del Or tdbesta.Tag = s_MdoData_Vis))
  ' Decimo Primera pestaña
  cmdActionterceros(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbterceros.Tag <> s_MdoData_Ins)
  cmdActionterceros(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbterceros.Tag = s_MdoData_Upd Or tdbterceros.Tag = s_MdoData_Vis))
  cmdActionterceros(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbterceros.Tag = s_MdoData_Del Or tdbterceros.Tag = s_MdoData_Vis))
  ' Decimo Segunda pestaña
  cmdActionsuspension(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbsuspension.Tag <> s_MdoData_Ins)
  cmdActionsuspension(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbsuspension.Tag = s_MdoData_Upd Or tdbsuspension.Tag = s_MdoData_Vis))
  cmdActionsuspension(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbsuspension.Tag = s_MdoData_Del Or tdbsuspension.Tag = s_MdoData_Vis))
  ' Decimo Tercera pestaña
  cmdActioncomprobantes(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbcomprobantes.Tag <> s_MdoData_Ins)
  cmdActioncomprobantes(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbcomprobantes.Tag = s_MdoData_Upd Or tdbcomprobantes.Tag = s_MdoData_Vis))
  cmdActioncomprobantes(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbcomprobantes.Tag = s_MdoData_Del Or tdbcomprobantes.Tag = s_MdoData_Vis))
  
End Sub
Sub RecuperarContrato()
  
  ' Genero la cadena de seleccion
  s_Sql = "SELECT CONCAT(numdocumen, '-', ano, mes, dia) AS numcontrato, CONCAT(numdocumen, ano, mes, dia) AS codcontrato,"
  s_Sql = s_Sql & " fechaini, fechafin, observacion, archivo, estadocon,tipcon, pltipcontrato.destco as destco"
  s_Sql = s_Sql & " FROM plcontrato"
  s_Sql = s_Sql & " inner join pltipcontrato on plcontrato.tipcon=pltipcontrato.codtco "
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " ORDER BY numcontrato"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbContrato.DataSource = porstRecordset
  
End Sub
Sub RecuperarEstudios()
  
  ' Genero la cadena de seleccion
  s_Sql = "SELECT est.institucion, est.orden, est.grado, plniveducativo.desniv as desniv, "
  s_Sql = s_Sql & "est.fechaini, est.fechafin, est.observacion "
  s_Sql = s_Sql & "FROM plestudios est "
  s_Sql = s_Sql & "inner join plniveducativo on est.grado=plniveducativo.codniv "
  s_Sql = s_Sql & "WHERE est.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND est.codpsn='" & txtCodigo & "' "
  s_Sql = s_Sql & "ORDER BY orden"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbEstudio.DataSource = porstRecordset
  
End Sub
Sub RecuperarExperiencia()
  
  ' Genero la cadena de seleccion
  s_Sql = "SELECT exl.empresa, exl.orden, exl.codcgo, cgo.descgo, "
  s_Sql = s_Sql & "exl.fechaini, exl.fechafin, exl.observacion "
  s_Sql = s_Sql & "FROM plexpelaboral exl "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON exl.codcls=cgo.codcls AND exl.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "WHERE exl.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND exl.codpsn='" & txtCodigo & "' "
  s_Sql = s_Sql & "ORDER BY orden"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbExperiencia.DataSource = porstRecordset
  
End Sub
Sub RecuperarFamiliares()
  
  ' Genero la cadena de seleccion
  s_Sql = "SELECT fam.orden, CONCAT(IFNULL(fam.apepaterno, ''), ' ', IFNULL(fam.apematerno, ''), ', ', IFNULL(fam.nombres, '')) AS nombresfam, "
  s_Sql = s_Sql & "fam.apepaterno, fam.apematerno, fam.nombres, fam.fecnacimiento, fam.sexofam, fam.coddci, "
  s_Sql = s_Sql & "dci.sigladci, fam.numdociden, fam.vinculo, fam.cartamed, fam.domicilio, fam.codvia, "
  s_Sql = s_Sql & "fam.nomviadom, fam.numerdom, fam.intedom, fam.codzona, fam.nomzonadom, fam.refedom, "
  s_Sql = s_Sql & "fam.ubigeodom, fam.incapacidad, fam.certificadomed, fam.motivoina, fam.estadofam,fam.tipdocpaternidad,fam.acrepaternidad,fam.fecalta,fam.fecbaja "
  s_Sql = s_Sql & "FROM plfamiliares fam "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON fam.coddci=dci.coddci "
  s_Sql = s_Sql & "WHERE fam.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND fam.codpsn='" & txtCodigo & "' "
  s_Sql = s_Sql & "ORDER BY orden"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbFamiliar.DataSource = porstRecordset
  
End Sub
Private Sub RecuperarRemuneraciones()
  Dim n_TotalRemunera As Double
  Dim s_Moneda As String * 1
  
  ' Inicializo el arreglo
  a_RemuneraDefa.ReDim 1, 0, 0, 4
  s_Moneda = IIf(chkPagoDolar.Value, "E", "N")
  ' Genero la cadena de seleccion
  s_Sql = "SELECT cxp.codcpc, cpc.descpc, codmon, IFNULL(rxd.imporemune, 0) AS imporemune,"
  s_Sql = s_Sql & " IF(cpc.tipocpc='0', 'Ingreso', IF(cpc.tipocpc='1','Descuento', 'Aporte')) AS destipocpc"
  s_Sql = s_Sql & " FROM ((plconceplanilla cxp"
  s_Sql = s_Sql & " LEFT JOIN plremudefa rxd USING(codcls, codcpc))"
  s_Sql = s_Sql & " LEFT JOIN plconcepto cpc USING(codcpc))"
  s_Sql = s_Sql & " WHERE cxp.codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND cxp.defaultcpc='" & s_Estado_Act & "'"
  s_Sql = s_Sql & " AND rxd.codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " UNION"
  s_Sql = s_Sql & " SELECT cxp.codcpc, cpc.descpc, '" & s_Moneda & "' AS codmon, 0 AS imporemune,"
  s_Sql = s_Sql & " IF(cpc.tipocpc='0', 'Ingreso', IF(cpc.tipocpc='1','Descuento', 'Aporte')) AS destipocpc"
  s_Sql = s_Sql & " FROM (plconceplanilla cxp"
  s_Sql = s_Sql & " LEFT JOIN plconcepto cpc USING(codcpc))"
  s_Sql = s_Sql & " WHERE cxp.codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND cxp.defaultcpc='" & s_Estado_Act & "'"
  s_Sql = s_Sql & " AND NOT EXISTS(SELECT * FROM plremudefa rxd"
  s_Sql = s_Sql & " WHERE rxd.codcls=cxp.codcls"
  s_Sql = s_Sql & " AND rxd.codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " AND rxd.codcpc=cxp.codcpc)"
  s_Sql = s_Sql & " ORDER BY codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  n_TotalRemunera = 0
  
  ' Si hay registros  de remuneraciones
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    a_RemuneraDefa.ReDim 1, porstRecordset.RecordCount, 0, 4
    n_Cuonter = 0
    While Not porstRecordset.EOF
      n_Cuonter = n_Cuonter + 1
      a_RemuneraDefa(n_Cuonter, 0) = gdl_Funcion.aTexto(porstRecordset!codcpc)
      a_RemuneraDefa(n_Cuonter, 1) = gdl_Funcion.aTexto(porstRecordset!descpc)
      a_RemuneraDefa(n_Cuonter, 2) = gdl_Funcion.aTexto(porstRecordset!destipocpc)
      a_RemuneraDefa(n_Cuonter, 3) = gdl_Funcion.aTexto(porstRecordset!codmon)
      a_RemuneraDefa(n_Cuonter, 4) = CDec(porstRecordset!imporemune)
      n_TotalRemunera = n_TotalRemunera + CDec(porstRecordset!imporemune)
      porstRecordset.MoveNext
    Wend
  End If
  ' Cierro el recordset y saco del entorno
  porstRecordset.Close: Set porstRecordset = Nothing
  
  ' Asigno el arreglo a la grilla y relleno la misma
  Set tdbRegistro.Array = a_RemuneraDefa
  tdbRegistro.Rebind
  ' Visualizo total de remuneraciones
  lblTotalRemunera = FormatNumber(CDec(n_TotalRemunera), 2)
  
End Sub
Sub RecuperarRemDsctoAnterior()
  
  ' Genero la cadena de seleccion
  s_Sql = "SELECT res.codcpc, cpc.descpc, res.codpdo, res.secuencia, res.pdoano, res.pdomes, res.codmon, res.importe_mn, res.importe_me, "
  s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Ina & "', 'Ingreso', IF(res.tipocpc='" & s_Estado_Act & "', 'Descuento', 'Aporte')) AS destipocpc "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpsn='" & txtCodigo & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & s_PeriodoRemAper & "' "
  s_Sql = s_Sql & "ORDER BY codpdo, codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbAnterior.DataSource = porstRecordset

End Sub
Sub RecuperarEmpleadores()

 ' Genero la cadena de seleccion
  s_Sql = "SELECT ruc,razons,orden "
  s_Sql = s_Sql & " FROM plempleadores"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " ORDER BY ruc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbempleador.DataSource = porstRecordset

End Sub
Sub Recuperarestalaboral()

 ' Genero la cadena de seleccion
  s_Sql = "SELECT ano,mes,ruc,codest,tasa,orden "
  s_Sql = s_Sql & " FROM plestalaboral"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " ORDER BY ano,mes"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbesta.DataSource = porstRecordset

End Sub
Sub Recuperarterceros()

 ' Genero la cadena de seleccion
  s_Sql = "SELECT mes,ano,ruc,sctrs,sctrp,codest,tasa,importe,orden "
  s_Sql = s_Sql & " FROM plterceros"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " ORDER BY ano,mes"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbterceros.DataSource = porstRecordset

End Sub
Sub Recuperarsuspension()

 ' Genero la cadena de seleccion
  s_Sql = "SELECT orden,numero,fecha,ejercicio,medio "
  s_Sql = s_Sql & " FROM plsuspensionct"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " ORDER BY fecha"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbsuspension.DataSource = porstRecordset

End Sub
Sub Recuperarcomprobantes()

 ' Genero la cadena de seleccion
  s_Sql = "SELECT orden,tipo,serie,numero,monto,fecemision,fecpago,retencion "
  s_Sql = s_Sql & " FROM plcomprobantect"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & txtCodigo & "'"
  s_Sql = s_Sql & " ORDER BY fecpago"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  ' Asigno el recordset a la grilla y relleno la misma
  Set tdbcomprobantes.DataSource = porstRecordset

End Sub
Sub ShowScreen()
    
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    ' Pestaña inicial
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codpsn.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoDocu, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!coddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!numdociden.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!numdocmil.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkExtanjero, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtNombres(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!apepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtNombres(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!apematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtNombres(2), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!nombres.DefinedSize
    gdl_Procedure.EditCombo "AT", cmbSexo, -1, Me.Tag, False
    gdl_Procedure.EditDTPicker "AT", dtpFecha, Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    lblEdad = Trim((Year(Date) - Year(dtpFecha)) - IIf((Format(Date, "mm-dd") > Format(dtpFecha, "mm-dd")), 0, 1)) & " Años"
    
    gdl_Procedure.EditCombo "AT", cmbEstadoCivil, -1, Me.Tag, False
    gdl_Procedure.EditText "AT", txtUbigeo(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!ubigeonac.DefinedSize
    gdl_Procedure.EditText "AT", txtNacional, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!nacionalidad.DefinedSize
    gdl_Procedure.EditText "AT", txtEssalud, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!nroessalud.DefinedSize
    gdl_Procedure.EditText "AT", txtHijos, CInt(0), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtDependientes, CInt(0), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditOptionCheck "AT", chkDsctoJudicial, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtPorJudicial, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtCuenta(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!coddeudor.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codacredor.DefinedSize
    gdl_Procedure.EditText "AT", txtcorreo, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!correoelect.DefinedSize
    ' Inicializo las fotografia
    imgFoto.Picture = LoadPicture()
    imgFoto.Refresh
    
    ' Primera pestaña
    gdl_Procedure.EditText "AT", txtTipoVia, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codvia.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreVia, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!nomviadirec.DefinedSize
    gdl_Procedure.EditText "AT", txtnumero(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!numerdirec.DefinedSize
    gdl_Procedure.EditText "AT", txtnumero(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!intedirec.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoZona, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codzona.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreZona, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!nomzondirec.DefinedSize
    gdl_Procedure.EditText "AT", txtReferencia, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!refedirec.DefinedSize
    gdl_Procedure.EditText "AT", txtLargaDistancia, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codldn.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!telefono.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!celular.DefinedSize
    gdl_Procedure.EditText "AT", txtUbigeo(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!ubigeodir.DefinedSize
  
    ' Segunda pestaña
    gdl_Procedure.EditDTPicker "AT", dtpFecIngreso, Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", chkReingreso, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtJornadaLabor, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtEntidadAfp, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codafp.DefinedSize
    gdl_Procedure.EditText "AT", txtNumeroAfp, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!numeroafp.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkComisMixta, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtTipoTraba, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codtpt.DefinedSize
    gdl_Procedure.EditText "AT", txtCargo, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codcgo.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkCargoConfi, False, Me.Tag, True
    gdl_Procedure.EditCombo "AT", cmbCargoConfi, 0, Me.Tag, False
    gdl_Procedure.EditText "AT", txtProfesion, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codpfs.DefinedSize
    gdl_Procedure.EditText "AT", txtCondicion, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codcdt.DefinedSize
    gdl_Procedure.EditText "AT", txtCenCosto, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtUbicacion, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codubica.DefinedSize
    gdl_Procedure.EditText "AT", txtSeccion, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codsec.DefinedSize
    
    gdl_Procedure.EditOptionCheck "AT", chkPagoDolar, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtTipPag, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!tippago.DefinedSize
    gdl_Procedure.EditText "AT", txtPeriodicidad, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!periodicidad.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbcopago.DefinedSize
    gdl_Procedure.EditText "AT", txtNroCuenta(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!cuentapago.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkInterbank(0), False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtBanco(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbnkpago.DefinedSize
    
    gdl_Procedure.EditOptionCheck "AT", chkCtsDolar, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkCtsDeposito, True, Me.Tag, True
    gdl_Procedure.EditText "AT", txtBanco(2), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbcocts.DefinedSize
    gdl_Procedure.EditText "AT", txtNroCuenta(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!cuentacts.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkInterbank(1), False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtBanco(3), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbnkcts.DefinedSize
    gdl_Procedure.EditText "AT", txtNroCuenta(2), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!cuentaibankcts.DefinedSize
    
    gdl_Procedure.EditText "AT", txtCodEps, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!codeps.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optRegimen(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optRegimen(1), False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optRegimen(2), False, Me.Tag, True
    gdl_Procedure.EditMask "AT", mskFecInicio, "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditOptionCheck "AT", chkEssaludVida, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkSindicato, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chk27252, False, Me.Tag, True
    gdl_Procedure.EditCombo "AT", cmbSituacion, 0, Me.Tag, False
    gdl_Procedure.EditMask "AT", mskFecSitua, Date, Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecBaja, "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditOptionCheck "AT", chkRL, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkDIS, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkMAX, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkREG, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkNOC, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkQUI, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkOIQ, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtSiteps, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!siteps.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkSM, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkMRF, False, Me.Tag, True
    gdl_Procedure.EditCombo "AT", cmbforprofesional, -1, Me.Tag, False
    gdl_Procedure.EditText "AT", txtfinperiodo, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!finperiodo.DefinedSize
    gdl_Procedure.EditText "AT", txtmodformativa, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!modformativa.DefinedSize
    
    gdl_Procedure.EditCombo "AT", cmbsctrs, -1, Me.Tag, False
    gdl_Procedure.EditCombo "AT", cmbsctrp, -1, Me.Tag, False
    
    gdl_Procedure.EditOptionCheck "AT", chkPE, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtcatocupacional, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!cmbcatocupacional.DefinedSize
    gdl_Procedure.EditText "AT", txttributacion, "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!cmbtributacion.DefinedSize
    
    ' Tercera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkRemImprecisa, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemIntegral(0), False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemIntegral(1), False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemIntegral(2), False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemuneNeta, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtConcepto(0), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!netocpc.DefinedSize
    gdl_Procedure.EditText "AT", txtConcepto(1), "", Me.Tag, False, fPersonal.dcaRegistro.Recordset!variacpc.DefinedSize
    gdl_Procedure.EditText "AT", txtRemuNeta, FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify

    ' Selecciono la pestaña inicial
    tasRegister.Tabs(1).Selected = True
  Else
    ' Pestaña inicial
    gdl_Procedure.EditText "PK", txtCodigo, fPersonal.dcaRegistro.Recordset!codpsn, Me.Tag, False, fPersonal.dcaRegistro.Recordset!codpsn.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoDocu, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!coddci), Me.Tag, False, fPersonal.dcaRegistro.Recordset!coddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!numdociden), Me.Tag, False, fPersonal.dcaRegistro.Recordset!numdociden.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!numdocmil), Me.Tag, False, fPersonal.dcaRegistro.Recordset!numdocmil.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkExtanjero, (fPersonal.dcaRegistro.Recordset!naciextrapsn = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtNombres(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!apepaterno), Me.Tag, False, fPersonal.dcaRegistro.Recordset!apepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtNombres(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!apematerno), Me.Tag, False, fPersonal.dcaRegistro.Recordset!apematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtNombres(2), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!nombres), Me.Tag, False, fPersonal.dcaRegistro.Recordset!nombres.DefinedSize
    gdl_Procedure.EditCombo "AT", cmbSexo, fPersonal.dcaRegistro.Recordset!sexopsn, Me.Tag, False
    gdl_Procedure.EditDTPicker "AT", dtpFecha, fPersonal.dcaRegistro.Recordset!fecnacimiento, Me.Tag, True, s_FormatoFecha, dtpShortDate
    lblEdad = Trim((Year(Date) - Year(dtpFecha)) - IIf((Format(Date, "mm-dd") > Format(dtpFecha, "mm-dd")), 0, 1)) & " Años"
    
    n_Cuonter = IIf(fPersonal.dcaRegistro.Recordset!estcivilpsn = "S", 0, IIf(fPersonal.dcaRegistro.Recordset!estcivilpsn = "C", 1, IIf(fPersonal.dcaRegistro.Recordset!estcivilpsn = "V", 2, IIf(fPersonal.dcaRegistro.Recordset!estcivilpsn = "D", 3, 4))))
    gdl_Procedure.EditCombo "AT", cmbEstadoCivil, n_Cuonter, Me.Tag, False
    gdl_Procedure.EditText "AT", txtUbigeo(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!ubigeonac), Me.Tag, False, fPersonal.dcaRegistro.Recordset!ubigeonac.DefinedSize
    gdl_Procedure.EditText "AT", txtNacional, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!nacionalidad), Me.Tag, False, fPersonal.dcaRegistro.Recordset!nacionalidad.DefinedSize
    gdl_Procedure.EditText "AT", txtEssalud, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!nroessalud), Me.Tag, False, fPersonal.dcaRegistro.Recordset!nroessalud.DefinedSize
    gdl_Procedure.EditText "AT", txtHijos, CInt(fPersonal.dcaRegistro.Recordset!numhijo), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtDependientes, CInt(fPersonal.dcaRegistro.Recordset!numdepen), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditOptionCheck "AT", chkDsctoJudicial, (fPersonal.dcaRegistro.Recordset!dctojudicial = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtPorJudicial, FormatNumber(fPersonal.dcaRegistro.Recordset!pordsctojudi, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtCuenta(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!coddeudor), Me.Tag, False, fPersonal.dcaRegistro.Recordset!coddeudor.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codacredor), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codacredor.DefinedSize
    gdl_Procedure.EditText "AT", txtcorreo, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!correoelect), Me.Tag, False, fPersonal.dcaRegistro.Recordset!correoelect.DefinedSize
    
    'Cargamos la Fotografia del personal
    s_Sql = "SELECT codpsn, fotopsn "
    s_Sql = s_Sql & "FROM plpersonal "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpsn='" & Trim(txtCodigo.Text) & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ReadImagen porstRecordset, imgFoto, "fotopsn"
    porstRecordset.Close
    
    ' Primera pestaña
    gdl_Procedure.EditText "AT", txtTipoVia, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codvia), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codvia.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreVia, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!nomviadirec), Me.Tag, False, fPersonal.dcaRegistro.Recordset!nomviadirec.DefinedSize
    gdl_Procedure.EditText "AT", txtnumero(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!numerdirec), Me.Tag, False, fPersonal.dcaRegistro.Recordset!numerdirec.DefinedSize
    gdl_Procedure.EditText "AT", txtnumero(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!intedirec), Me.Tag, False, fPersonal.dcaRegistro.Recordset!intedirec.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoZona, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codzona), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codzona.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreZona, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!nomzondirec), Me.Tag, False, fPersonal.dcaRegistro.Recordset!nomzondirec.DefinedSize
    gdl_Procedure.EditText "AT", txtReferencia, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!refedirec), Me.Tag, False, fPersonal.dcaRegistro.Recordset!refedirec.DefinedSize
    gdl_Procedure.EditText "AT", txtLargaDistancia, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codldn), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codldn.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!telefono), Me.Tag, False, fPersonal.dcaRegistro.Recordset!telefono.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!celular), Me.Tag, False, fPersonal.dcaRegistro.Recordset!celular.DefinedSize
    gdl_Procedure.EditText "AT", txtUbigeo(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!ubigeodir), Me.Tag, False, fPersonal.dcaRegistro.Recordset!ubigeodir.DefinedSize
    
    ' Segunda pestaña
    gdl_Procedure.EditDTPicker "AT", dtpFecIngreso, fPersonal.dcaRegistro.Recordset!fecingreso, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", chkReingreso, (fPersonal.dcaRegistro.Recordset!reingreso = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtJornadaLabor, FormatNumber(fPersonal.dcaRegistro.Recordset!jornadalaboral, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtEntidadAfp, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codafp), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codafp.DefinedSize
    gdl_Procedure.EditText "AT", txtNumeroAfp, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!numeroafp), Me.Tag, False, fPersonal.dcaRegistro.Recordset!numeroafp.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkComisMixta, (fPersonal.dcaRegistro.Recordset!afpmixta = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtTipoTraba, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codtpt), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codtpt.DefinedSize
    gdl_Procedure.EditText "AT", txtCargo, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codcgo), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codcgo.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkCargoConfi, (fPersonal.dcaRegistro.Recordset!cgoconfianza <> s_Estado_Ina), Me.Tag, True
    n_Cuonter = CInt(fPersonal.dcaRegistro.Recordset!cgoconfianza)
    gdl_Procedure.EditCombo "AT", cmbCargoConfi, n_Cuonter, Me.Tag, False
    gdl_Procedure.EditText "AT", txtProfesion, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codpfs), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codpfs.DefinedSize
    gdl_Procedure.EditText "AT", txtCondicion, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codcdt), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codcdt.DefinedSize
    gdl_Procedure.EditText "AT", txtCenCosto, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codcco), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtUbicacion, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codubica), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codubica.DefinedSize
    gdl_Procedure.EditText "AT", txtSeccion, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codsec), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codsec.DefinedSize
    
    gdl_Procedure.EditOptionCheck "AT", chkPagoDolar, (fPersonal.dcaRegistro.Recordset!pagodolar = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtPeriodicidad, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!periodicidad), Me.Tag, False, fPersonal.dcaRegistro.Recordset!periodicidad.DefinedSize
    gdl_Procedure.EditText "AT", txtTipPag, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!tippago), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbcopago.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codbcopago), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbcopago.DefinedSize
    gdl_Procedure.EditText "AT", txtNroCuenta(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!cuentapago), Me.Tag, False, fPersonal.dcaRegistro.Recordset!cuentapago.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkInterbank(0), (fPersonal.dcaRegistro.Recordset!interbankpago = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtBanco(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codbnkpago), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbnkpago.DefinedSize
    
    gdl_Procedure.EditOptionCheck "AT", chkCtsDolar, (fPersonal.dcaRegistro.Recordset!ctsdolar = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkCtsDeposito, (fPersonal.dcaRegistro.Recordset!ctsdeposito = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtBanco(2), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codbcocts), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbcocts.DefinedSize
    gdl_Procedure.EditText "AT", txtNroCuenta(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!cuentacts), Me.Tag, False, fPersonal.dcaRegistro.Recordset!cuentacts.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkInterbank(1), (fPersonal.dcaRegistro.Recordset!interbankcts = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtBanco(3), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codbnkcts), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbnkcts.DefinedSize
    gdl_Procedure.EditText "AT", txtNroCuenta(2), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!cuentaibankcts), Me.Tag, False, fPersonal.dcaRegistro.Recordset!cuentaibankcts.DefinedSize
    
    gdl_Procedure.EditText "AT", txtCodEps, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!codeps), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codeps.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optRegimen(0), (fPersonal.dcaRegistro.Recordset!regpension = s_Estado_Ina), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optRegimen(1), (fPersonal.dcaRegistro.Recordset!regpension = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optRegimen(2), (fPersonal.dcaRegistro.Recordset!regpension = s_Estado_Blq), Me.Tag, True
    gdl_Procedure.EditMask "AT", mskFecInicio, IIf(IsNull(fPersonal.dcaRegistro.Recordset!fecingregpen), "", fPersonal.dcaRegistro.Recordset!fecingregpen), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditOptionCheck "AT", chkEssaludVida, (fPersonal.dcaRegistro.Recordset!essvida = s_Estado_Act), Me.Tag, True
    'gdl_Procedure.EditOptionCheck "AT", chkSctr, (fPersonal.dcaRegistro.Recordset!cobsctr = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkSindicato, (fPersonal.dcaRegistro.Recordset!afilsindical = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chk27252, (fPersonal.dcaRegistro.Recordset!chk27252 = s_Estado_Act), Me.Tag, True
    n_Cuonter = IIf(fPersonal.dcaRegistro.Recordset!estadopsn = "A", 0, IIf(fPersonal.dcaRegistro.Recordset!estadopsn = "V", 1, IIf(fPersonal.dcaRegistro.Recordset!estadopsn = "L", 2, IIf(fPersonal.dcaRegistro.Recordset!estadopsn = "N", 3, IIf(fPersonal.dcaRegistro.Recordset!estadopsn = "P", 4, 5)))))
    gdl_Procedure.EditCombo "AT", cmbSituacion, n_Cuonter, Me.Tag, False
    gdl_Procedure.EditMask "AT", mskFecSitua, IIf(IsNull(fPersonal.dcaRegistro.Recordset!fecestado), "", fPersonal.dcaRegistro.Recordset!fecestado), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecBaja, IIf(IsNull(fPersonal.dcaRegistro.Recordset!fecbaja), "", fPersonal.dcaRegistro.Recordset!fecbaja), Me.Tag, True, "##/##/####"
    
    'gdl_Procedure.EditOptionCheck "AT", chkSCTRP, (fPersonal.dcaRegistro.Recordset!chkSCTRP = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRL, (fPersonal.dcaRegistro.Recordset!chkRL = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkDIS, (fPersonal.dcaRegistro.Recordset!chkDIS = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkMAX, (fPersonal.dcaRegistro.Recordset!chkMAX = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkREG, (fPersonal.dcaRegistro.Recordset!chkREG = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkNOC, (fPersonal.dcaRegistro.Recordset!chkNOC = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkQUI, (fPersonal.dcaRegistro.Recordset!chkQUI = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkOIQ, (fPersonal.dcaRegistro.Recordset!chkOIQ = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtSiteps, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!siteps), Me.Tag, False, fPersonal.dcaRegistro.Recordset!codbcopago.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", chkSM, (fPersonal.dcaRegistro.Recordset!segmedico = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkMRF, (fPersonal.dcaRegistro.Recordset!resfamiliar = s_Estado_Act), Me.Tag, True
    
    n_Cuonter = IIf(fPersonal.dcaRegistro.Recordset!forprofesional = "1", 0, IIf(fPersonal.dcaRegistro.Recordset!forprofesional = "2", 1, IIf(fPersonal.dcaRegistro.Recordset!forprofesional = "3", 2, IIf(fPersonal.dcaRegistro.Recordset!forprofesional = "4", 3, 3))))
    gdl_Procedure.EditCombo "AT", cmbforprofesional, n_Cuonter, Me.Tag, False
    
    gdl_Procedure.EditText "AT", txtfinperiodo, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!finperiodo), Me.Tag, False, fPersonal.dcaRegistro.Recordset!finperiodo.DefinedSize
    gdl_Procedure.EditText "AT", txtmodformativa, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!modformativa), Me.Tag, False, fPersonal.dcaRegistro.Recordset!modformativa.DefinedSize
    
    n_Cuonter = IIf(fPersonal.dcaRegistro.Recordset!cobsctr = "0", 0, IIf(fPersonal.dcaRegistro.Recordset!cobsctr = "1", 1, 2))
    gdl_Procedure.EditCombo "AT", cmbsctrs, n_Cuonter, Me.Tag, False
    
    n_Cuonter = IIf(fPersonal.dcaRegistro.Recordset!chkSCTRP = "0", 0, IIf(fPersonal.dcaRegistro.Recordset!chkSCTRP = "1", 1, 2))
    gdl_Procedure.EditCombo "AT", cmbsctrp, n_Cuonter, Me.Tag, False
    
    gdl_Procedure.EditOptionCheck "AT", chkPE, (fPersonal.dcaRegistro.Recordset!chkPE = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtcatocupacional, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!cmbcatocupacional), Me.Tag, False, fPersonal.dcaRegistro.Recordset!cmbcatocupacional.DefinedSize
    gdl_Procedure.EditText "AT", txttributacion, gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!cmbtributacion), Me.Tag, False, fPersonal.dcaRegistro.Recordset!cmbtributacion.DefinedSize
    
    ' Tercera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkEssaludVida, (fPersonal.dcaRegistro.Recordset!essvida = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemImprecisa, (fPersonal.dcaRegistro.Recordset!remimprecisa = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemIntegral(0), (fPersonal.dcaRegistro.Recordset!remintegralgrati = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemIntegral(1), (fPersonal.dcaRegistro.Recordset!remintegralvaca = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemIntegral(2), (fPersonal.dcaRegistro.Recordset!remintegralcts = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkRemuneNeta, (fPersonal.dcaRegistro.Recordset!remuneta = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtConcepto(0), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!netocpc), Me.Tag, False, fPersonal.dcaRegistro.Recordset!netocpc.DefinedSize
    gdl_Procedure.EditText "AT", txtConcepto(1), gdl_Funcion.aTexto(fPersonal.dcaRegistro.Recordset!variacpc), Me.Tag, False, fPersonal.dcaRegistro.Recordset!variacpc.DefinedSize
    gdl_Procedure.EditText "AT", txtRemuNeta, FormatNumber(fPersonal.dcaRegistro.Recordset!imporemuneto, 2), Me.Tag, False, 18, vbRightJustify
  End If
  ' Pestaña Inicial
  imgFoto.Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoDocu, "DI")
  lblUbigeo(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUbigeo(0), "UG")
  lblHelp(15) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtNacional, "NA")
   
  ' Primera pestaña
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoVia, "TV")
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoZona, "TZ")
  lblUbigeo(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUbigeo(1), "UG")
  lblHelp(26) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtLargaDistancia.Text, "LD")
  ' Segunda pestaña
  lblTiempo = Trim(Year(Date) - Year(dtpFecIngreso)) & " Año(s) de Servicio"
  lblHelp(3).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoTraba, "TT")
  lblHelp(4).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtCargo, "DC")
  lblHelp(5).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtProfesion, "PF")
  lblHelp(6).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, ps_CodEmpresa, txtCenCosto, "CC")
  lblHelp(7).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtCodEps, "ES")
  lblHelp(8).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtEntidadAfp, "EP")
  lblHelp(9).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(0), "EB")
  lblHelp(24).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(1), "EB")
  lblHelp(10).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(2), "EB")
  lblHelp(25).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(3), "EB")
  lblHelp(11).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtUbicacion, "UL")
  lblHelp(12).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtSeccion, "SE")
  lblHelp(13).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtConcepto(0), "CP")
  lblHelp(14).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtConcepto(1), "CP")
  lblHelp(16).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtSiteps, "SI")
  lblHelp(17).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipPag, "PA")
  lblHelp(18).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtfinperiodo, "FP")
  lblHelp(19).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtmodformativa, "MF")
  lblHelp(20).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtPeriodicidad, "PE")
  lblHelp(21).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtcatocupacional, "CY")
  lblHelp(22).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txttributacion, "CZ")
  lblHelp(23).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtCondicion.Text, "ST")
  
  ' Tercera pestaña
  lblNombre(0).Caption = "Nombres : " & Trim(txtNombres(0)) & " " & Trim(txtNombres(1)) & "; " & Trim(txtNombres(2)) & "  "
  RecuperarRemuneraciones
  ' Cuarta pestaña
  lblNombre(1).Caption = lblNombre(0).Caption
  RecuperarExperiencia
  ' Quinta pestaña
  lblNombre(2).Caption = lblNombre(0).Caption
  RecuperarEstudios
  ' Sexta pestaña
  lblNombre(3).Caption = lblNombre(0).Caption
  RecuperarContrato
  ' Septima pestaña
  lblNombre(4).Caption = lblNombre(0).Caption
  RecuperarFamiliares
  ' Octava pestaña
  lblNombre(5).Caption = lblNombre(0).Caption
  RecuperarRemDsctoAnterior
 ' Novena pestaña
  lblNombre(7).Caption = lblNombre(0).Caption
  RecuperarEmpleadores
 ' Decima pestaña
  lblNombre(8).Caption = lblNombre(0).Caption
  Recuperarestalaboral
 ' Decima primera pestaña
  lblNombre(9).Caption = lblNombre(0).Caption
  Recuperarterceros
  ' Decima segunda pestaña
  lblNombre(10).Caption = lblNombre(0).Caption
  Recuperarsuspension
  ' Decima tercera pestaña
  lblNombre(11).Caption = lblNombre(0).Caption
  Recuperarcomprobantes

End Sub
']
Private Sub chkCtsDeposito_Click(Value As Integer)
  If Value Then chkRemIntegral(2).Value = vbUnchecked
End Sub
Private Sub chkDsctoJudicial_Click(Value As Integer)
  If Value = vbUnchecked Then
    txtPorJudicial.Text = FormatNumber(0, 2)
  End If
End Sub
Private Sub chkRemImprecisa_Click(Value As Integer)
  If Value Then chkRemuneNeta.Value = vbUnchecked
End Sub
Private Sub chkRemIntegral_Click(Index As Integer, Value As Integer)
  If (Value And Index = 2) Then chkCtsDeposito.Value = vbUnchecked
End Sub
Private Sub chkRemuneNeta_Click(Value As Integer)
  
  If Value = vbUnchecked Then
    txtConcepto(0) = ""
    lblHelp(13) = ""
    txtRemuNeta = FormatNumber(0, 2)
    txtConcepto(1) = ""
    lblHelp(14) = ""
  Else
    chkRemImprecisa.Value = vbUnchecked
  End If

End Sub
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ' Verifico nivel de usuario
  If (Index = 0 And ps_NivelUsr = NivelUsuario.Auxiliar) Then Beep: MsgBox "Usuario no puede actualizar Información " & lblTitle, vbExclamation: Me.Tag = s_MdoData_Vis: Index = 3
  tdbExperiencia.Tag = s_MdoData_Vis: tdbEstudio.Tag = s_MdoData_Vis
  tdbContrato.Tag = s_MdoData_Vis: tdbFamiliar.Tag = s_MdoData_Vis
  tdbAnterior.Tag = s_MdoData_Vis: tdbesta.Tag = s_MdoData_Vis
  tdbempleador.Tag = s_MdoData_Vis: tdbterceros.Tag = s_MdoData_Vis
  tdbsuspension.Tag = s_MdoData_Vis: tdbcomprobantes.Tag = s_MdoData_Vis
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
    If n_IndexTabs = 0 Then
      'txtTipoDocu.SetFocus
    ElseIf n_IndexTabs = 1 Then
      txtTipoVia.SetFocus
    ElseIf n_IndexTabs = 2 Then
      'txtTipoTraba.SetFocus
    ElseIf n_IndexTabs = 3 Then
      'tdbRegistro.SetFocus
    ElseIf n_IndexTabs = 4 Then
      cmdActionExp(0).SetFocus
    ElseIf n_IndexTabs = 5 Then
      cmdActionEst(0).SetFocus
    ElseIf n_IndexTabs = 6 Then
      cmdActionCon(0).SetFocus
    ElseIf n_IndexTabs = 7 Then
      cmdActionFam(0).SetFocus
    ElseIf n_IndexTabs = 8 Then
      cmdActionAnt(0).SetFocus
    ElseIf n_IndexTabs = 9 Then
      cmdActionEmp(0).SetFocus
    ElseIf n_IndexTabs = 10 Then
      cmdActionEsta(0).SetFocus
    ElseIf n_IndexTabs = 11 Then
      cmdActionterceros(0).SetFocus
    ElseIf n_IndexTabs = 12 Then
      cmdActionsuspension(0).SetFocus
    ElseIf n_IndexTabs = 13 Then
      cmdActioncomprobantes(0).SetFocus
    End If
  End If
  If Index <> 1 Then Exit Sub
    
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim(txtNombres(0)) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Personal = Trim(txtCodigo)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpsn")
    a_Valores = Array(ps_ClsPlanilla, s_Personal)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plpersonal", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fPersonal.dcaRegistro, fPersonal.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fPersonal.dcaRegistro.Recordset.EOF And fPersonal.dcaRegistro.Recordset.BOF) Or fPersonal.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fPersonal.dcaRegistro.Recordset.Find ("codpsn >= '" & s_Personal & "'")
      If fPersonal.dcaRegistro.Recordset.EOF Then fPersonal.dcaRegistro.Recordset.MoveLast
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
Private Sub cmdActionAnt_Click(Index As Integer)

  cmdActionAnt(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbAnterior.Tag <> s_MdoData_Ins)
  cmdActionAnt(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbAnterior.Tag = s_MdoData_Upd Or tdbAnterior.Tag = s_MdoData_Vis))
  cmdActionAnt(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbAnterior.Tag = s_MdoData_Del Or tdbAnterior.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbAnterior.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbAnterior.Tag = s_MdoData_Ins
    fAbcRemunAnterior.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbAnterior.VisibleRows = 0) Then
      tdbAnterior.Tag = Choose(Index, s_MdoData_Del, s_MdoData_Vis)
      fAbcRemunAnterior.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionAnt(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionAnt(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionAnt(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionCon_Click(Index As Integer)
  
  cmdActionCon(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbContrato.Tag <> s_MdoData_Ins)
  cmdActionCon(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbContrato.Tag = s_MdoData_Upd Or tdbContrato.Tag = s_MdoData_Vis))
  cmdActionCon(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbContrato.Tag = s_MdoData_Del Or tdbContrato.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbContrato.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbContrato.Tag = s_MdoData_Ins
    fAbcContrato.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbContrato.VisibleRows = 0) Then
      tdbContrato.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcContrato.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionCon(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionCon(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionCon(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionEmp_Click(Index As Integer)
  cmdActionEmp(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbempleador.Tag <> s_MdoData_Ins)
  cmdActionEmp(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbempleador.Tag = s_MdoData_Upd Or tdbempleador.Tag = s_MdoData_Vis))
  cmdActionEmp(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbempleador.Tag = s_MdoData_Del Or tdbempleador.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbempleador.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbempleador.Tag = s_MdoData_Ins
    fAbcEmpleadores.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbempleador.VisibleRows = 0) Then
      tdbempleador.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcEmpleadores.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionEmp(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionEmp(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionEmp(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionEsta_Click(Index As Integer)
  
  cmdActionEsta(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbesta.Tag <> s_MdoData_Ins)
  cmdActionEsta(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbesta.Tag = s_MdoData_Upd Or tdbesta.Tag = s_MdoData_Vis))
  cmdActionEsta(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbesta.Tag = s_MdoData_Del Or tdbesta.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbesta.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbesta.Tag = s_MdoData_Ins
    fAbcEstaLaboral.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbesta.VisibleRows = 0) Then
      tdbesta.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcEstaLaboral.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionEsta(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionEsta(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionEsta(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionterceros_Click(Index As Integer)
  
  cmdActionterceros(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbterceros.Tag <> s_MdoData_Ins)
  cmdActionterceros(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbterceros.Tag = s_MdoData_Upd Or tdbterceros.Tag = s_MdoData_Vis))
  cmdActionterceros(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbterceros.Tag = s_MdoData_Del Or tdbterceros.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbterceros.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbterceros.Tag = s_MdoData_Ins
    fAbcTerceros.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbterceros.VisibleRows = 0) Then
      tdbterceros.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcTerceros.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionterceros(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionterceros(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionterceros(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub

Private Sub cmdActionEst_Click(Index As Integer)
  
  cmdActionEst(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbEstudio.Tag <> s_MdoData_Ins)
  cmdActionEst(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbEstudio.Tag = s_MdoData_Upd Or tdbEstudio.Tag = s_MdoData_Vis))
  cmdActionEst(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbEstudio.Tag = s_MdoData_Del Or tdbEstudio.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbEstudio.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbEstudio.Tag = s_MdoData_Ins
    fAbcEstudioRealizado.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbEstudio.VisibleRows = 0) Then
      tdbEstudio.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcEstudioRealizado.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionEst(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionEst(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionEst(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionExp_Click(Index As Integer)
  
  cmdActionExp(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbExperiencia.Tag <> s_MdoData_Ins)
  cmdActionExp(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbExperiencia.Tag = s_MdoData_Upd Or tdbExperiencia.Tag = s_MdoData_Vis))
  cmdActionExp(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbExperiencia.Tag = s_MdoData_Del Or tdbExperiencia.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbExperiencia.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbExperiencia.Tag = s_MdoData_Ins
    fAbcExperienciaLaboral.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbExperiencia.VisibleRows = 0) Then
      tdbExperiencia.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcExperienciaLaboral.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionExp(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionExp(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionExp(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionFam_Click(Index As Integer)

  cmdActionFam(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbFamiliar.Tag <> s_MdoData_Ins)
  cmdActionFam(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbFamiliar.Tag = s_MdoData_Upd Or tdbFamiliar.Tag = s_MdoData_Vis))
  cmdActionFam(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbFamiliar.Tag = s_MdoData_Del Or tdbFamiliar.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbFamiliar.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbFamiliar.Tag = s_MdoData_Ins
    fAbcFamiliar.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbFamiliar.VisibleRows = 0) Then
      tdbFamiliar.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcFamiliar.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionFam(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionFam(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionFam(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActionsuspension_Click(Index As Integer)

  cmdActionsuspension(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbsuspension.Tag <> s_MdoData_Ins)
  cmdActionsuspension(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbsuspension.Tag = s_MdoData_Upd Or tdbsuspension.Tag = s_MdoData_Vis))
  cmdActionsuspension(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbsuspension.Tag = s_MdoData_Del Or tdbsuspension.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbsuspension.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbsuspension.Tag = s_MdoData_Ins
    fAbcSuspension.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbsuspension.VisibleRows = 0) Then
      tdbsuspension.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcSuspension.Show
    Else
      ' Habili/inhabilito los controles
      cmdActionsuspension(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionsuspension(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActionsuspension(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdActioncomprobantes_Click(Index As Integer)

  cmdActioncomprobantes(0).Enabled = (Me.Tag = s_MdoData_Upd And tdbcomprobantes.Tag <> s_MdoData_Ins)
  cmdActioncomprobantes(1).Enabled = (Me.Tag = s_MdoData_Upd And (tdbcomprobantes.Tag = s_MdoData_Upd Or tdbcomprobantes.Tag = s_MdoData_Vis))
  cmdActioncomprobantes(2).Enabled = (Me.Tag = s_MdoData_Upd And (tdbcomprobantes.Tag = s_MdoData_Del Or tdbcomprobantes.Tag = s_MdoData_Vis))
  ' Inicializo el modo de registro o selección
  tdbcomprobantes.Tag = ""
  Select Case Index
   Case 0 ' Nuevo registro
    tdbcomprobantes.Tag = s_MdoData_Ins
    fAbcComprobantes.Show
   Case 1, 2  ' Modificar, eliminar registro
    If Not (tdbcomprobantes.VisibleRows = 0) Then
      tdbcomprobantes.Tag = IIf(Index = 1, s_MdoData_Del, s_MdoData_Vis)
      fAbcComprobantes.Show
    Else
      ' Habili/inhabilito los controles
      cmdActioncomprobantes(0).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActioncomprobantes(1).Enabled = (Me.Tag = s_MdoData_Upd)
      cmdActioncomprobantes(2).Enabled = (Me.Tag = s_MdoData_Upd)
    End If
  End Select

End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim s_TablaHelp As String
  
  s_SqlHelp = ""
  If n_IndexHelp = Index And Index <> 1 Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If
  ReDim aElemento(2, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Código", "Descripción")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "coddci", "desdci")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 734.7402, 3465.071)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0)
  Next n_Cuonter
  
  Select Case Index
   Case 0     ' Tipo de documento de identidad
    aElemento(0, 1) = "coddci": aElemento(1, 1) = "desdci"
    s_TablaHelp = "Documentos de Identidad"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("dci", aElemento(0, 1), "", "")
   Case 1     ' Tipo de via de dirección
    aElemento(0, 1) = "codvia": aElemento(1, 1) = "desvia"
    s_TablaHelp = "Tipos de Via"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("via", aElemento(0, 1), "", "")
   Case 2     ' Tipo de zona de direccion
    aElemento(0, 1) = "codzona": aElemento(1, 1) = "deszona"
    s_TablaHelp = "Tipos de Zona"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("zon", aElemento(0, 1), "", "")
   Case 3     ' Tipo de trabajador
    aElemento(0, 1) = "codtpt": aElemento(1, 1) = "destpt"
    s_TablaHelp = "Tipo de Trabajador"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("tpt", aElemento(0, 1), "", "")
   Case 4     ' Cargo de personal
    aElemento(0, 1) = "codcgo": aElemento(1, 1) = "descgo"
    s_TablaHelp = "Cargo de  Personal"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cgo", aElemento(0, 1), ps_ClsPlanilla, "")
   Case 5     ' Profesión de personal
    aElemento(0, 1) = "codpfs": aElemento(1, 1) = "despfs"
    s_TablaHelp = "Profesión de Personal"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("pfs", aElemento(0, 1), "", "")
   Case 6     ' Centro de costo
    aElemento(0, 1) = "codcco": aElemento(1, 1) = "detcco"
    s_TablaHelp = "Centro de Costos"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cco", aElemento(0, 1), pn_NivelCenCosto, "")
   Case 7     ' Entidad de servicios
    aElemento(0, 1) = "codeps": aElemento(1, 1) = "deseps"
    s_TablaHelp = "Entidad de EPS"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("eps", aElemento(0, 1), "", "")
   Case 8     ' Entidad de pensión
    aElemento(0, 1) = "codafp": aElemento(1, 1) = "desafp"
    s_TablaHelp = "Entidad de Pensión"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("afp", aElemento(0, 1), "", "")
   Case 9, 10, 24, 25  ' Entidad bancaria
    aElemento(0, 1) = "codbco": aElemento(1, 1) = "desbco"
    s_TablaHelp = "Entidad Bancaria"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("bco", aElemento(0, 1), "", "")
   Case 11     ' Ubicación o localidad de agencia
    aElemento(0, 1) = "codubica": aElemento(1, 1) = "desubica"
    s_TablaHelp = "Ubicación o Localidad"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ubi", aElemento(0, 1), "", "")
   Case 12     ' Sección de la empresa
    aElemento(0, 1) = "codsec": aElemento(1, 1) = "dessec"
    s_TablaHelp = "Sección de la Empresa"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("sec", aElemento(0, 1), "", "")
   Case 13, 14    ' Conceptos de planilla
    aElemento(0, 1) = "codcpc": aElemento(1, 1) = "descpc"
    s_TablaHelp = "Concepto de Planilla"
    s_Personal = ps_ClsPlanilla & IIf(Index = 13, "F", "C") & s_Estado_Ina
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", aElemento(0, 1), s_Personal, "")
    Case 15     ' Nacionalidad
    aElemento(0, 1) = "codnac": aElemento(1, 1) = "desnac"
    s_TablaHelp = "Nacionalidad"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("nac", aElemento(0, 1), "", "")
    Case 16     ' Situacion Trabajdor
    aElemento(0, 1) = "codstp": aElemento(1, 1) = "desstp"
    s_TablaHelp = "Situacion Trabajador"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("stp", aElemento(0, 1), "", "")
    Case 17     ' Tipo de Pago
    aElemento(0, 1) = "codtip": aElemento(1, 1) = "destip"
    s_TablaHelp = "Tipo de Pago"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("tip", aElemento(0, 1), "", "")
    Case 18     ' Motivo de Fin de Periodo
    aElemento(0, 1) = "codmof": aElemento(1, 1) = "desmof"
    s_TablaHelp = "Motivo de Fin de Periodo"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("mof", aElemento(0, 1), "", "")
    Case 19     ' Modalidad Formativa
    aElemento(0, 1) = "codmfo": aElemento(1, 1) = "desmfo"
    s_TablaHelp = "Modalidad Formativa"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("mfo", aElemento(0, 1), "", "")
    Case 20     ' Periodicidad
    aElemento(0, 1) = "codprd": aElemento(1, 1) = "desprd"
    s_TablaHelp = "Periodicidad"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("prd", aElemento(0, 1), "", "")
   Case 21     ' categoria ocupacional
    aElemento(0, 1) = "codcao": aElemento(1, 1) = "descao"
    s_TablaHelp = "catocupacional"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cao", aElemento(0, 1), "", "")
   Case 22     ' tributacion
    aElemento(0, 1) = "codctr": aElemento(1, 1) = "desctr"
    s_TablaHelp = "tributacion"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ctr", aElemento(0, 1), "", "")
   Case 23     ' Condición de trabajo
    aElemento(0, 1) = "codcdt": aElemento(1, 1) = "descdt"
    s_TablaHelp = "Condición de Trabajo"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cdt", aElemento(0, 1), ps_ClsPlanilla, "")
   Case 26     ' Larga distancia
    aElemento(0, 1) = "codldn": aElemento(1, 1) = "desldn"
    s_TablaHelp = "Código Larga Distancia"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ldn", aElemento(0, 1), ps_ClsPlanilla, "")
  End Select
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Cuonter, 0) = ""
    aElementos(n_Cuonter, 1) = n_BackColorHelp#: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbHelp, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbHelp, s_TablaHelp, 2
  ' Recupera información
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & IIf(Index = 6, ps_DaBasCon, ps_DataBase), dcaHelp, tdbHelp, s_Sql, adCmdText, adLockReadOnly
  
  ' Muestra la grilla de ayuda
  n_Cuonter = n_Cuonter + (frmRegister.Top + tasRegister.Top + frmOpciones(n_IndexTabs).Top) + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  n_Cuonter = IIf((Index >= 13 And Index < 24), 6000, n_Cuonter)
  tdbHelp.Top = IIf(n_Cuonter > 4400, 4100, n_Cuonter)
  n_Cuonter = tasRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Left = IIf(Index >= 13, 2450, IIf(Index >= 9, 4050, n_Cuonter))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fPersonal.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fPersonal.dcaRegistro.Recordset.BOF Then fPersonal.dcaRegistro.Recordset.MovePrevious
           If fPersonal.dcaRegistro.Recordset.BOF Then fPersonal.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fPersonal.dcaRegistro.Recordset.EOF Then fPersonal.dcaRegistro.Recordset.MoveNext
           If fPersonal.dcaRegistro.Recordset.EOF Then fPersonal.dcaRegistro.Recordset.MoveLast
   Case 3: fPersonal.dcaRegistro.Recordset.MoveLast
  End Select

lblTitle = "Trabajador(a) " & " " & Trim(txtNombres(0)) & " " & Trim(txtNombres(1)) & " " & Trim(txtNombres(2))

End Sub

Private Sub cmdUbigeo_Click(Index As Integer)
  Set o_SwSelUbica = fAbcPersonal: n_SwSelUbica = Index
  fSeleccionUbigeo.Show vbModal
  Set o_SwSelUbica = Nothing
  Exit Sub
End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1, s_Extranjero As String * 1, s_EstadoCivil As String * 1
  Dim s_DsctoJudicial As String * 1, s_CargoConfianza As String * 1
  Dim s_PagoDolar As String * 1, s_CtsDolar As String * 1, s_CtsDeposito As String * 1
  Dim s_RemuIntegralGrati As String * 1, s_RemuIntegralVaca As String * 1
  Dim s_RemuIntegralCts As String * 1, s_RemuneNeta As String * 1, s_RemuneImprecisa As String * 1
  Dim s_EssaludVida As String * 1, s_CoberturaScrt As String * 1
  Dim s_Sindical As String * 1, s_RegPension As String * 1, s_Reingreso As String * 1
  Dim s_ComisionMixta As String * 1

  Dim s_chkSCTRP As String * 1, s_chkRL As String * 1
  Dim s_chkDIS As String * 1, s_chkMAX As String * 1
  Dim s_chkREG As String * 1, s_chkNOC As String * 1
  Dim s_chkQUI As String * 1, s_chkOIQ As String * 1
  Dim s_siteps As String * 1, s_tippago As String * 1
  Dim s_chkSM As String * 1, s_chkMRF As String * 1
  Dim s_forprofesional As String * 1, s_finperiodo As String * 1
  Dim s_modformativa As String * 1, s_27252 As String * 1
  Dim s_InterbankPago As String * 1, s_InterbankCts As String * 1

  Dim s_sctrs As String * 1, s_sctrp As String * 1
  Dim s_chkPE As String * 1
  
  Dim n_Expresion As Long
  
  ' Realizo las validaciones de los campos a actualizar
  ' Pestaña Inicial
  If txtCodigo.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: txtCodigo.SetFocus: Exit Sub
  If txtTipoDocu.Text = "" Then Beep: MsgBox "Debe Ingresar el Tipo de Documento " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: txtTipoDocu.SetFocus: Exit Sub
  If lblHelp(0).Caption = "???" Then Beep: MsgBox "Tipo Documento Identidad no es valido; Verificar", vbExclamation: tasRegister.Tabs(1).Selected = True: txtTipoDocu.SetFocus: Exit Sub
  If txtDocumento(0).Text = "" Then Beep: MsgBox "Debe Ingresar el Numero de Documento " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: txtDocumento(0).SetFocus: Exit Sub
  If Trim(txtNombres(0).Text) = "" Or Trim(txtNombres(1)) = "" Or Trim(txtNombres(2)) = "" Then Beep: MsgBox "Debe Ingresar los nombres del trabajador", vbExclamation: tasRegister.Tabs(1).Selected = True: txtNombres(0).SetFocus: Exit Sub
  If cmbSexo.Text = "" Then Beep: MsgBox "Seleccione Sexo " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: cmbSexo.SetFocus: Exit Sub
  If cmbEstadoCivil.Text = "" Then Beep: MsgBox "Seleccione Estado Civil " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: cmbEstadoCivil.SetFocus: Exit Sub
  s_EstadoCivil = Choose(cmbEstadoCivil.ListIndex + 1, "S", "C", "V", "D", "O")
  If txtNacional.Text = "" Then Beep: MsgBox "Debe Ingresar la Nacionalidad del " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: txtNacional.SetFocus: Exit Sub
  If lblHelp(15).Caption = "???" Then Beep: MsgBox "Nacionalidad no valida; Verificar ", vbExclamation: tasRegister.Tabs(1).Selected = True: txtNacional.SetFocus: Exit Sub
  If txtEntidadAfp.Text = "" Then Beep: MsgBox "Debe Ingresar la Entidad Administradora de Pensiones del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtEntidadAfp.SetFocus: Exit Sub
  If lblHelp(8).Caption = "???" Then Beep: MsgBox "Entidad de Pensiones no valida; Verificar ", vbExclamation: tasRegister.Tabs(3).Selected = True: txtEntidadAfp.SetFocus: Exit Sub
  
  If txtTipoTraba.Text = "" Then Beep: MsgBox "Debe Ingresar el Tipo" & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtTipoTraba.SetFocus: Exit Sub
  If lblHelp(3).Caption = "???" Then Beep: MsgBox "Tipo no valido; Verificar ", vbExclamation: tasRegister.Tabs(3).Selected = True: txtTipoTraba.SetFocus: Exit Sub
  If txtCargo.Text = "" Then Beep: MsgBox "Debe Ingresar el Cargo del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtCargo.SetFocus: Exit Sub
  If lblHelp(4).Caption = "???" Then Beep: MsgBox "Cargo no valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtCargo.SetFocus: Exit Sub
  If lblHelp(23).Caption = "???" Then Beep: MsgBox "Condición de Trabajo no valida; Verificar ", vbExclamation: tasRegister.Tabs(3).Selected = True: txtCondicion.SetFocus: Exit Sub
  If txtProfesion.Text = "" Then Beep: MsgBox "Debe Ingresar la Profesion-Ocupacion del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtProfesion.SetFocus: Exit Sub
  If lblHelp(5).Caption = "???" Then Beep: MsgBox "Profesion-Ocupacion no valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtProfesion.SetFocus: Exit Sub
  If txtPeriodicidad.Text = "" Then Beep: MsgBox "Debe Ingresar la Periodicidad del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtPeriodicidad.SetFocus: Exit Sub
  If lblHelp(20).Caption = "???" Then Beep: MsgBox "Tipo de Periodicidad no valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtPeriodicidad.SetFocus: Exit Sub
  If txtTipPag.Text = "" Then Beep: MsgBox "Debe Ingresar el Tipo de Pago del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtTipPag.SetFocus: Exit Sub
  If lblHelp(17).Caption = "???" Then Beep: MsgBox "Tipo de Pago no valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtTipPag.SetFocus: Exit Sub
  If txtCodEps.Text = "" Then Beep: MsgBox "Debe Ingresar la Entidad de Eps del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtCodEps.SetFocus: Exit Sub
  If lblHelp(7).Caption = "???" Then Beep: MsgBox "Entidad Eps no valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtCodEps.SetFocus: Exit Sub
  If txtSiteps.Text = "" Then Beep: MsgBox "Debe Ingresar la Situacion de Eps del " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: txtSiteps.SetFocus: Exit Sub
  If lblHelp(16).Caption = "???" Then Beep: MsgBox "Situacion Eps no valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtSiteps.SetFocus: Exit Sub
  If cmbforprofesional.Text = "" Then Beep: MsgBox "Seleccione Formacion Profesional " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: cmbforprofesional.SetFocus: Exit Sub
  s_Extranjero = IIf(chkExtanjero.Value, s_Estado_Act, s_Estado_Ina)
  s_DsctoJudicial = IIf(chkDsctoJudicial.Value, s_Estado_Act, s_Estado_Ina)
  If lblUbigeo(0).Caption = "???" Then Beep: MsgBox "Ubicacion Geografica de Nacimiento no Valido; Verificar", vbExclamation: tasRegister.Tabs(1).Selected = True: txtUbigeo(0).SetFocus: Exit Sub
  
  If cmbsctrs.Text = "" Then Beep: MsgBox "Seleccione Seguro Complementario de Trabajo de Riesgo Salud " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: cmbsctrs.SetFocus: Exit Sub
  If cmbsctrp.Text = "" Then Beep: MsgBox "Seleccione Seguro Complementario de Trabajo de Riesgo Pension " & lblTitle, vbExclamation: tasRegister.Tabs(3).Selected = True: cmbsctrp.SetFocus: Exit Sub
  
  ' Primera pestaña
  If lblHelp(1).Caption = "???" Then Beep: MsgBox "Dirección - Tipo de Via no es valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtTipoVia.SetFocus: Exit Sub
  If lblHelp(2).Caption = "???" Then Beep: MsgBox "Dirección - Tipo de Zona no es valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtTipoZona.SetFocus: Exit Sub
  If lblUbigeo(1).Caption = "???" Then Beep: MsgBox "Dirección - Ubicacion Geografica no es valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtUbigeo(1).SetFocus: Exit Sub
  If lblHelp(26).Caption = "???" Then Beep: MsgBox "Dirección - Codigo Larga Distancia Nacional; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtLargaDistancia.SetFocus: Exit Sub
  If ((txtTelefono(0).Text <> "" Or txtTelefono(1).Text <> "") And txtLargaDistancia.Text = "") Then Beep: MsgBox "Registrar Codigo Larga Distancia Nacional; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtLargaDistancia.SetFocus: Exit Sub
  
  ' Segunda pestaña
  s_CargoConfianza = IIf(chkCargoConfi.Value, IIf(cmbCargoConfi.ListIndex = 5, cmbCargoConfi.ListIndex, IIf(cmbCargoConfi.ListIndex = 4, cmbCargoConfi.ListIndex, IIf(cmbCargoConfi.ListIndex = 3, cmbCargoConfi.ListIndex, IIf(cmbCargoConfi.ListIndex = 2, s_Estado_Blq, s_Estado_Act)))), s_Estado_Ina)
  s_PagoDolar = IIf(chkPagoDolar.Value, s_Estado_Act, s_Estado_Ina)
  s_InterbankPago = IIf(chkInterbank(0).Value, s_Estado_Act, s_Estado_Ina)
  s_InterbankCts = IIf(chkInterbank(1).Value, s_Estado_Act, s_Estado_Ina)
  s_CtsDeposito = IIf(chkCtsDeposito.Value, s_Estado_Act, s_Estado_Ina)
  s_CtsDolar = IIf(chkCtsDolar.Value, s_Estado_Act, s_Estado_Ina)
  s_Reingreso = IIf(chkReingreso.Value, s_Estado_Act, s_Estado_Ina)
  
  s_ComisionMixta = IIf(chkComisMixta.Value, s_Estado_Act, s_Estado_Ina)
  s_RegPension = IIf(optRegimen(0).Value, s_Estado_Ina, IIf(optRegimen(1).Value, s_Estado_Act, s_Estado_Blq))
  s_EssaludVida = IIf(chkEssaludVida.Value, s_Estado_Act, s_Estado_Ina)
  's_CoberturaScrt = IIf(chkSctr.Value, s_Estado_Act, s_Estado_Ina)
  s_Sindical = IIf(chkSindicato.Value, s_Estado_Act, s_Estado_Ina)
  s_27252 = IIf(chk27252.Value, s_Estado_Act, s_Estado_Ina)
  s_Estado = Choose(cmbSituacion.ListIndex + 1, "A", "V", "L", "N", "P", "I")
  s_forprofesional = Choose(cmbforprofesional.ListIndex + 1, "1", "2", "3", "4")
  
  s_sctrs = Choose(cmbsctrs.ListIndex + 1, "0", "1", "2")
  s_sctrp = Choose(cmbsctrp.ListIndex + 1, "0", "1", "2")
  
  If lblHelp(3).Caption = "???" Then Beep: MsgBox "Tipo de Trabajador no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtTipoTraba.SetFocus: Exit Sub
  If lblHelp(4).Caption = "???" Then Beep: MsgBox "Cargo de Personal no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtCargo.SetFocus: Exit Sub
  If lblHelp(5).Caption = "???" Then Beep: MsgBox "Profesión no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtProfesion.SetFocus: Exit Sub
  If txtCenCosto.Text = "" Then Beep: MsgBox "Centro de Costo no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtCenCosto.SetFocus: Exit Sub
  If lblHelp(6).Caption = "???" Then Beep: MsgBox "Centro de Costo no es valido; Verificar", vbExclamation: txtCenCosto.SetFocus: Exit Sub
  If lblHelp(7).Caption = "???" Then Beep: MsgBox "Entidad Prestadora de Servicio no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtCodEps.SetFocus: Exit Sub
  If txtEntidadAfp.Text = "" Then: Beep: MsgBox "Entidad de Pensiones - APF no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtEntidadAfp.SetFocus: Exit Sub
  If lblHelp(8).Caption = "???" Then Beep: MsgBox "Entidad de Pensiones - APF no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtEntidadAfp.SetFocus: Exit Sub
  If lblHelp(9).Caption = "???" Then Beep: MsgBox "Entidad Bancaria de Pago Remuneraciones no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtBanco(0).SetFocus: Exit Sub
  If (chkInterbank(0).Value And txtBanco(1).Text = "") Then Beep: MsgBox "Entidad Interbancaria de Pago Remuneraciones no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtBanco(1).SetFocus: Exit Sub
  If lblHelp(24).Caption = "???" Then Beep: MsgBox "Entidad Interbancaria de Pago Remuneraciones no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtBanco(1).SetFocus: Exit Sub
  If lblHelp(10).Caption = "???" Then Beep: MsgBox "Entidad Bancaria de Deposito CTS no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtBanco(2).SetFocus: Exit Sub
  If (chkInterbank(1).Value And txtBanco(3).Text = "") Then Beep: MsgBox "Entidad Interbancaria de Deposito CTS no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtBanco(3).SetFocus: Exit Sub
  If lblHelp(25).Caption = "???" Then Beep: MsgBox "Entidad Interbancaria de Deposito CTS no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtBanco(3).SetFocus: Exit Sub
  If (chkInterbank(1).Value And txtNroCuenta(2).Text = "") Then Beep: MsgBox "Cuenta Interbancaria de Deposito CTS no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtNroCuenta(2).SetFocus: Exit Sub

  If txtUbicacion.Text = "" Then: Beep: MsgBox "Ubicación o Localidad no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtUbicacion.SetFocus: Exit Sub
  If lblHelp(11).Caption = "???" Then Beep: MsgBox "Ubicación o Localidad no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtUbicacion.SetFocus: Exit Sub
  If txtSeccion.Text = "" Then: Beep: MsgBox "Sección de la Empresa no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtSeccion.SetFocus: Exit Sub
  If lblHelp(12).Caption = "???" Then Beep: MsgBox "Sección de la Empresa no es valido; Verificar", vbExclamation: tasRegister.Tabs(3).Selected = True: txtSeccion.SetFocus: Exit Sub
  
  If mskFecInicio.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecInicio, 1900) Then tasRegister.Tabs(3).Selected = True: mskFecInicio.SetFocus: Exit Sub
  End If
  If mskFecSitua.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecSitua, 1900) Then tasRegister.Tabs(3).Selected = True: mskFecSitua.SetFocus: Exit Sub
  End If
  If mskFecBaja.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecBaja, 1900) Then tasRegister.Tabs(3).Selected = True: mskFecBaja.SetFocus: Exit Sub
  End If

  's_chkSCTRP = IIf(chkSCTRP.Value, s_Estado_Act, s_Estado_Ina)
  s_chkRL = IIf(chkRL.Value, s_Estado_Act, s_Estado_Ina)
  s_chkDIS = IIf(chkDIS.Value, s_Estado_Act, s_Estado_Ina)
  s_chkMAX = IIf(chkMAX.Value, s_Estado_Act, s_Estado_Ina)
  s_chkREG = IIf(chkREG.Value, s_Estado_Act, s_Estado_Ina)
  s_chkNOC = IIf(chkNOC.Value, s_Estado_Act, s_Estado_Ina)
  s_chkQUI = IIf(chkQUI.Value, s_Estado_Act, s_Estado_Ina)
  s_chkOIQ = IIf(chkOIQ.Value, s_Estado_Act, s_Estado_Ina)
  s_chkSM = IIf(chkSM.Value, s_Estado_Act, s_Estado_Ina)
  s_chkMRF = IIf(chkMRF.Value, s_Estado_Act, s_Estado_Ina)
  s_chkPE = IIf(chkPE.Value, s_Estado_Act, s_Estado_Ina)

  ' Cuarta pestaña
  s_RemuIntegralGrati = IIf(chkRemIntegral(0).Value, s_Estado_Act, s_Estado_Ina)
  s_RemuIntegralVaca = IIf(chkRemIntegral(1).Value, s_Estado_Act, s_Estado_Ina)
  s_RemuIntegralCts = IIf(chkRemIntegral(2).Value, s_Estado_Act, s_Estado_Ina)
  s_RemuneNeta = IIf(chkRemuneNeta.Value, s_Estado_Act, s_Estado_Ina)
  s_RemuneImprecisa = IIf(chkRemImprecisa.Value, s_Estado_Act, s_Estado_Ina)
  
  n_Expresion = IIf(ps_NivelUsr = NivelUsuario.Auxiliar, 1, 4)
  If (s_RemuneImprecisa = s_Estado_Act And s_RemuneNeta = s_Estado_Act) Then
    Beep
    MsgBox "Selecione Remuneración Neta o Imprecisa", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then chkRemImprecisa.SetFocus
    Exit Sub
  End If
  If (s_RemuIntegralCts = s_Estado_Act And s_CtsDeposito = s_Estado_Act) Then
    Beep
    MsgBox "Selecione Remuneración Integral o Deposito de C.T.S.", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then chkRemIntegral(2).SetFocus
    Exit Sub
  End If
  txtConcepto(0).Text = IIf(s_RemuneNeta = s_Estado_Ina, "", txtConcepto(0).Text)
  txtRemuNeta.Text = FormatNumber(IIf(s_RemuneNeta = s_Estado_Ina, 0, txtRemuNeta), 2)
  txtConcepto(1).Text = IIf(s_RemuneNeta = s_Estado_Ina, "", txtConcepto(1).Text)
  If (s_RemuneNeta = s_Estado_Act And txtConcepto(0).Text = "") Then
    Beep
    MsgBox "Ingrese concepto de Remuneración Neta", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then txtConcepto(0).SetFocus
    Exit Sub
  End If
  If (s_RemuneNeta = s_Estado_Act And Val(txtRemuNeta.Text) <= 0) Then
    Beep
    MsgBox "Ingrese importe de Remuneración Neta", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then txtRemuNeta.SetFocus
    Exit Sub
  End If
  If (s_RemuneNeta = s_Estado_Act And txtConcepto(1).Text = "") Then
    Beep
    MsgBox "Ingrese concepto de Ajuste de Remenración", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then txtConcepto(1).SetFocus
    Exit Sub
  End If
  If (s_RemuneNeta = s_Estado_Act And lblHelp(13).Caption = "???") Then
    Beep
    MsgBox "Concepto de Remeneración Neta no Valido; Verificar", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then txtConcepto(0).SetFocus
    Exit Sub
  End If
  If (s_RemuneNeta = s_Estado_Act And lblHelp(14).Caption = "???") Then
    Beep
    MsgBox "Concepto de Regularización de Remenración no Valido; Verificar", vbExclamation: tasRegister.Tabs(n_Expresion).Selected = True
    If n_Expresion = 4 Then txtConcepto(1).SetFocus
    Exit Sub
  End If
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Personal = txtCodigo
  
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpsn", "apepaterno", "apematerno", "nombres", "fecnacimiento", "ubigeonac", "nacionalidad", "naciextrapsn", "sexopsn", "estcivilpsn", "coddci", "numdociden", "numdocmil", "numhijo", "numdepen", "dctojudicial", "pordsctojudi", "coddeudor", "codacredor", _
             "codvia", "nomviadirec", "numerdirec", "intedirec", "codzona", "nomzondirec", "refedirec", "codldn", "telefono", "celular", "ubigeodir", _
             "fecingreso", "reingreso", "jornadalaboral", "codtpt", "codcgo", "cgoconfianza", "codpfs", "codcdt", "codcco", "codafp", "numeroafp", "afpmixta", "pagodolar", "periodicidad", "tippago", "codbcopago", "cuentapago", "interbankpago", "codbnkpago", "ctsdeposito", "ctsdolar", "codbcocts", "cuentacts", "interbankcts", "codbnkcts", "cuentaibankcts", _
             "codeps", "regpension", "fecingregpen", "essvida", "cobsctr", "afilsindical", "remintegralgrati", "remintegralvaca", "remintegralcts", "remimprecisa ", "remuneta", "netocpc", "variacpc", "imporemuneto", _
             "fecestado", "fecbaja", "nroessalud", "codubica", "codsec", "estadopsn", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"), "correoelect", "chkSCTRP", "chkRL", "chkDIS", "chkMAX", "chkREG", "chkNOC", "chkQUI", "chkOIQ", "siteps", "segmedico", "resfamiliar", "forprofesional", "finperiodo", "modformativa", "chkpe", "cmbcatocupacional", "cmbtributacion", "chk27252")
  a_Valores = Array(ps_ClsPlanilla, txtCodigo, Trim(txtNombres(0)), Trim(txtNombres(1)), Trim(txtNombres(2)), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtUbigeo(0)), Trim(txtNacional), s_Extranjero, Trim(cmbSexo.ListIndex), s_EstadoCivil, Trim(txtTipoDocu), Trim(txtDocumento(0)), Trim(txtDocumento(1)), CInt(txtHijos), CInt(txtDependientes), s_DsctoJudicial, CDec(txtPorJudicial), Trim(txtCuenta(0)), Trim(txtCuenta(1)), _
              Trim(txtTipoVia), Trim(txtNombreVia), Trim(txtnumero(0)), Trim(txtnumero(1)), Trim(txtTipoZona), Trim(txtNombreZona), Trim(txtReferencia.Text), Trim(txtLargaDistancia.Text), Trim(txtTelefono(0)), Trim(txtTelefono(1)), Trim(txtUbigeo(1)), _
              Format(dtpFecIngreso, s_FmtFechMysql_0), s_Reingreso, CDec(txtJornadaLabor.Text), Trim(txtTipoTraba), Trim(txtCargo), s_CargoConfianza, Trim(txtProfesion), Trim(txtCondicion.Text), Trim(txtCenCosto), Trim(txtEntidadAfp), Trim(txtNumeroAfp), s_ComisionMixta, s_PagoDolar, Trim(txtPeriodicidad.Text), Trim(txtTipPag.Text), Trim(txtBanco(0).Text), Trim(txtNroCuenta(0)), s_InterbankPago, Trim(txtBanco(1).Text), s_CtsDeposito, s_CtsDolar, Trim(txtBanco(2).Text), Trim(txtNroCuenta(1).Text), s_InterbankCts, Trim(txtBanco(3).Text), Trim(txtNroCuenta(2).Text), _
              Trim(txtCodEps), s_RegPension, Format(mskFecInicio, s_FmtFechMysql_0), s_EssaludVida, s_sctrs, s_Sindical, s_RemuIntegralGrati, s_RemuIntegralVaca, s_RemuIntegralCts, s_RemuneImprecisa, s_RemuneNeta, Trim(txtConcepto(0)), Trim(txtConcepto(1)), CDec(txtRemuNeta), _
              Format(mskFecSitua, s_FmtFechMysql_0), Format(mskFecBaja, s_FmtFechMysql_0), Trim(txtEssalud), Trim(txtUbicacion), Trim(txtSeccion), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0), Trim(txtcorreo.Text), s_sctrp, s_chkRL, s_chkDIS, s_chkMAX, s_chkREG, s_chkNOC, s_chkQUI, s_chkOIQ, Trim(txtSiteps), s_chkSM, s_chkMRF, s_forprofesional, Trim(txtfinperiodo), Trim(txtmodformativa), s_chkPE, Trim(txtcatocupacional), Trim(txttributacion), s_27252)
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, _
             TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
             TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
             TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, _
             TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpsn")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
     
    If Not Records_Ins("plpersonal", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plpersonal", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  ' Realizo la grabacion de la imagen - foto
  s_Sql = "SELECT codcls, codpsn, fotopsn "
  s_Sql = s_Sql & "FROM plpersonal "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpsn='" & s_Personal & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenKeyset, adLockOptimistic, adUseClient, s_Sql)
  If Not WriteImagen(porstRecordset, imgFoto, "fotopsn") Then GoTo Error
  porstRecordset.Close
  
  '[ Realizo la grabacion de las remuneraciones
  ' Elimino las remuneraciones existentes
  s_Sql = "DELETE FROM plremudefa"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND codpsn='" & s_Personal & "'"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
  ' Fuerzo a que se actualice la grilla y refresco
  tdbRegistro.Update
  tdbRegistro.Refresh
  ' Grabo las nuevas remuneraciones
  For n_Cuonter = a_RemuneraDefa.LowerBound(1) To a_RemuneraDefa.UpperBound(1)
    If CDec(a_RemuneraDefa(n_Cuonter, 4)) <> 0 Then
      a_Campos = Array("codcls", "codpsn", "codcpc", "codmon", "imporemune", "usrcre", "fyhcre")
      a_Valores = Array(ps_ClsPlanilla, txtCodigo, Trim(a_RemuneraDefa(n_Cuonter, 0)), Trim(a_RemuneraDefa(n_Cuonter, 3)), CDec(a_RemuneraDefa(n_Cuonter, 4)), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
      If Not Records_Ins("plremudefa", a_Campos, a_Valores, a_Tipos) Then GoTo Error
    End If
  Next n_Cuonter
  ']
  
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fPersonal.dcaRegistro, fPersonal.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fPersonal.dcaRegistro.Recordset.Find ("codpsn='" & s_Personal & "'")
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
Private Sub dtpFecha_LostFocus()
  lblEdad = Trim(Year(Date) - Year(dtpFecha)) & " Años"
End Sub
Private Sub dtpFecIngreso_LostFocus()
  lblTiempo = Trim(Year(Date) - Year(dtpFecIngreso)) & " Año(s) de Servicio"
  If Me.Tag = s_MdoData_Ins Then mskFecSitua = Format(dtpFecIngreso.Value, s_FormatoFecha)
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

Set cnn = New ADODB.Connection
cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
cnn.CursorLocation = adUseClient
cnn.Open

  Dim Item As New ValueItem   ' Cambio el formato de la grilla columna de valores
    
  'Establece posición y titulo del formulario
  Me.Height = 9830: Me.Width = 9300
  Me.Left = 1080: Me.Top = 50
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Trabajadores"
  
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operación del registro
  Me.Tag = fPersonal.Tag
  tdbExperiencia.Tag = s_MdoData_Vis
  tdbEstudio.Tag = s_MdoData_Vis
  tdbContrato.Tag = s_MdoData_Vis
  tdbFamiliar.Tag = s_MdoData_Vis
  tdbAnterior.Tag = s_MdoData_Vis
  tdbempleador.Tag = s_MdoData_Vis
  tdbesta.Tag = s_MdoData_Vis
  tdbterceros.Tag = s_MdoData_Vis
  tdbsuspension.Tag = s_MdoData_Vis
  tdbcomprobantes.Tag = s_MdoData_Vis
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Cuonter = 0 To 2
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "anadir", "borrar", "modifica")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Cuonter
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo los graficos a los controles del toolbar adicionales
  For n_Cuonter = 0 To 2
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "anadir", "borrar", "modifica")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, "Añadir ", "Eliminar ", "Modificar ") & "Registro"
  Next n_Cuonter
  gdl_Procedure.ViewGrafics Me, cmdActionExp, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionEst, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionCon, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionFam, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionAnt, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionEmp, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionEsta, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionterceros, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActionsuspension, aElemento
  gdl_Procedure.ViewGrafics Me, cmdActioncomprobantes, aElemento
  
  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Cuonter = 0 To 3
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Cuonter
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fPersonal.dcaRegistro.Recordset.EOF Or fPersonal.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fPersonal.dcaRegistro.Recordset!codpsn
  
  ' Configuro los listados, datos adicionales
  For n_Cuonter = 0 To 1: cmbSexo.AddItem Choose(n_Cuonter + 1, "Masculino", "Femenino"): Next n_Cuonter
  For n_Cuonter = 0 To 4: cmbEstadoCivil.AddItem Choose(n_Cuonter + 1, "Soltero(a)", "Casado(a)", "Viudo(a)", "Divorciado(a)", "Conviviente"): Next n_Cuonter
  For n_Cuonter = 0 To 5: cmbSituacion.AddItem Choose(n_Cuonter + 1, "Activo", "Vacaciones", "Licencia", "Pre Natal", "Post Natal", "Inactivo"): Next n_Cuonter
  For n_Cuonter = 0 To 8: cmbCargoConfi.AddItem Choose(n_Cuonter + 1, "Ninguna", "Dirección - Presencial", "Confianza - Presencial", "Dirección - Teletrabajo Mixto", "Confianza - Teletrabajo Mixto", "Dirección - Teletrabajo Completo", "Confianza - Teletrabajo Completo", "Teletrabajo Mixto", "Teletrabajo Completo"): Next n_Cuonter
  For n_Cuonter = 0 To 3: cmbforprofesional.AddItem Choose(n_Cuonter + 1, "Centro Educativo", "Universidad", "Instituto", "Otros"): Next n_Cuonter
  
  For n_Cuonter = 0 To 2: cmbsctrs.AddItem Choose(n_Cuonter + 1, "Ninguno", "Essalud", "EPS"): Next n_Cuonter
  For n_Cuonter = 0 To 2: cmbsctrp.AddItem Choose(n_Cuonter + 1, "Ninguno", "ONP", "Seguro Privado"): Next n_Cuonter
  
  ' Asigno el control de datos  ala grilla
  tdbHelp.DataSource = dcaHelp
  
   '[ Configuración de la grilla de remuneración default
  ReDim aElemento(5, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Concepto", "Descripción", "Tipo", "Mon", "Importe")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "codcpc", "descpc", "destipocpc", "codmon", "imporemune")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 915.26, 3280.906, 900, 450, 1200)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbRightJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "", "", "standard")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 0)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, "Remuneraciones Default", 3
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(3).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(3).ValueItems.Validate = True
  tdbRegistro.Columns(3).ValueItems.Translate = True
  tdbRegistro.Columns(3).ValueItems.CycleOnClick = True
  For n_Cuonter = 0 To 1
    tdbRegistro.Columns(3).ValueItems.Add Item
    tdbRegistro.Columns(3).ValueItems.Item(n_Cuonter).Value = Choose(n_Cuonter + 1, s_Codmon_mn, s_Codmon_me)
    tdbRegistro.Columns(3).ValueItems.Item(n_Cuonter).DisplayValue = Choose(n_Cuonter + 1, s_Codmon_mn_Txt, s_Codmon_me_Txt)
  Next n_Cuonter
  ']
  
  '[ Configuración de la grilla de experiencia laboral
  ReDim aElemento(7, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Empresa", "Cargo", "Inicio", "Fin", "cargo", "observación", "orden")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "empresa", "descgo", "fechaini", "fechafin", "codcgo", "observacion", "orden")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 2520, 1800, 970, 970, 10, 10, 10)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", s_FormatoFecha, s_FormatoFecha, "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 1, 1, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbExperiencia, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbExperiencia, "Experiencia Laboral", 3
  ']
  
  '[ Configuración de la grilla de estudios realizados
  ReDim aElemento(7, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Institucion", "Grado", "Inicio", "Fin", "observación", "orden", "Grado")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "institucion", "grado", "fechaini", "fechafin", "observacion", "orden", "desniv")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 2520, 0, 970, 970, 10, 10, 3500)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", s_FormatoFecha, s_FormatoFecha, "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 1, 0, 0, 1, 1, 0)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbEstudio, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbEstudio, "Estudios Realizados", 3
  ']
  
  '[ Configuración de la grilla de contratos
  ReDim aElemento(8, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Número", "Inicio", "Fin", "Observación", "Ok", "Archivo", "Tipo", "Tipo")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "numcontrato", "fechaini", "fechafin", "observacion", "estadocon", "archivo", "tipcon", "destco")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 1650, 970, 970, 2370, 300, 2000, 0, 2500)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", s_FormatoFecha, s_FormatoFecha, "", "", "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 0, 1, 1, 0)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbContrato, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbContrato, "Contratos de Trabajo", 3
  
  tdbContrato.Columns(4).ValueItems.Presentation = dbgNormal
  tdbContrato.Columns(4).ValueItems.Translate = True
  For n_Cuonter = 0 To 1
    tdbContrato.Columns(4).ValueItems.Add Item
    tdbContrato.Columns(4).ValueItems.Item(n_Cuonter).Value = Choose(n_Cuonter + 1, s_Estado_Act, s_Estado_Ina)
    tdbContrato.Columns(4).ValueItems.Item(n_Cuonter).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Cuonter + 1, "estadok", "estadnok") & ".bmp")
  Next n_Cuonter
  ']
  
  '[ Configuración de la grilla de empleadores
  ReDim aElemento(3, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Ruc", "Razon Social", "Orden")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "ruc", "razons", "orden")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 1500, 3500, 0)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbempleador, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbempleador, "Otros Empleadores", 3
 
  ']
  
  '[ Configuración de la grilla de Empresas donde labora el Trabajador (Destaque)
 ReDim aElemento(6, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Año", "Mes", "Ruc", "Codigo Establecimiento", "Tasa", "orden")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "ano", "mes", "ruc", "codest", "tasa", "orden")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 1000, 1000, 1200, 2500, 600, 0)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbCenter, vbCenter)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "", "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 0, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbesta, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbesta, "Empresas donde labora el Trabajador (Destaque)", 3
  ']
  
  '[ Configuración de la grilla de Personal de Terceros
 ReDim aElemento(9, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Mes", "Año", "Ruc", "SCTR Salud", "SCTR Pension", "Cod. Establecimiento", "Tasa", "Importe", "Orden")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "mes", "ano", "ruc", "sctrs", "sctrp", "codest", "tasa", "importe", "orden")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 500, 500, 1200, 1200, 1200, 1800, 1000, 1000, 0)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbCenter, vbCenter, vbCenter, vbCenter, vbCenter)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 0, 0, 0, 0, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbterceros, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbterceros, "Personal de Terceros", 3
  ']
  
  
  '[ Configuración de la grilla de familiares
  ReDim aElemento(29, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Apellidos y Nombres", "Fecha Nac", "TDI", "Documento", "Vinculo", "Ok", "orden", "apepaterno", "apematerno", "nombres", "sexofam", "coddci", "cartamed", "domicilio", "codvia", "nomviadom", "numerdom", "intedom", "codzona", "nomzonadom", "refedom", "ubigeodom", "incapacidad", "certificadomed", "motivoina", "tipdocpaternidad", "acrepaternidad", "fecalta", "fecbaja")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "nombresfam", "fecnacimiento", "sigladci", "numdociden", "vinculo", "estadofam", "orden", "apepaterno", "apematerno", "nombres", "sexofam", "coddci", "cartamed", "domicilio", "codvia", "nomviadom", "numerdom", "intedom", "codzona", "nomzonadom", "refedom", "ubigeodom", "incapacidad", "certificadomed", "motivoina", "tipdocpaternidad", "acrepaternidad", "fecalta", "fecbaja")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 2790, 970, 350, 880, 980, 300, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 0, 0, 0, 0)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbCenter, vbLeftJustify, vbLeftJustify, vbCenter, vbRightJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", s_FormatoFecha, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
  Next n_Cuonter
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbFamiliar, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbFamiliar, "Datos Familiares", 3
  
  ' Cambio el formato de la grilla columna de valores
  tdbFamiliar.Columns(4).ValueItems.Presentation = dbgNormal
  tdbFamiliar.Columns(4).ValueItems.Translate = True
  For n_Cuonter = 0 To 4
    tdbFamiliar.Columns(4).ValueItems.Add Item
    tdbFamiliar.Columns(4).ValueItems.Item(n_Cuonter).Value = Trim(n_Cuonter)
    tdbFamiliar.Columns(4).ValueItems.Item(n_Cuonter).DisplayValue = Choose(n_Cuonter + 1, "Otro", "Hijo", "Conyuge", "Concubina(o)", "Gestante")
  Next n_Cuonter
  tdbFamiliar.Columns(5).ValueItems.Presentation = dbgNormal
  tdbFamiliar.Columns(5).ValueItems.Translate = True
  For n_Cuonter = 0 To 1
    tdbFamiliar.Columns(5).ValueItems.Add Item
    tdbFamiliar.Columns(5).ValueItems.Item(n_Cuonter).Value = Choose(n_Cuonter + 1, s_Estado_Ina, s_Estado_Act)
    tdbFamiliar.Columns(5).ValueItems.Item(n_Cuonter).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Cuonter + 1, "estadnok", "estadok") & ".bmp")
  Next n_Cuonter
  ']
  
  '[ Configuración de la grilla de ingresos y descuentos anteriores
  ReDim aElemento(8, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Mes", "Codigo", "Descripción", "Tipo", "Importe MN", "Importe ME", "Secuencia", "Moneda")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "pdomes", "codcpc", "descpc", "destipocpc", "importe_mn", "importe_me", "secuencia", "codmon")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 500, 700, 2030, 650, 1200, 1200, 10, 10)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "", "", "standard", "standard", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, dbgMergeRestricted, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgCenter, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 0, 0, 0, 0, 0, 0, 1, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbAnterior, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbAnterior, "Remuneración - Descuento Anterior", 3
  ']
  
  '[ Configuración de la grilla de Suspension de Cuarta Categoria
 ReDim aElemento(5, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Orden", "Numero", "Fecha", "Ejercicio", "Medio")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "orden", "numero", "fecha", "ejercicio", "medio")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 0, 1200, 1200, 1500, 0)
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbCenter)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 1, 0, 0, 0, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbsuspension, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbsuspension, "Suspension de Cuarta Categoria", 3
  ']
  
  '[ Configuración de la grilla de Comprobantes de Cuarta Categoria
 ReDim aElemento(8, 10)
  For n_Cuonter = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Cuonter, 0) = Choose(n_Cuonter + 1, "Orden", "Tipo", "Serie", "Numero", "Monto", "F. Emision", "F. Pago", "Retencion")
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "orden", "tipo", "serie", "numero", "monto", "fecemision", "fecpago", "retencion")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, 0, 1200, 1200, 1200, 1200, "1500", "1500", "0")
    aElemento(n_Cuonter, 3) = Choose(n_Cuonter + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbCenter, vbLeftJustify, vbLeftJustify, vbLeftJustify)
    aElemento(n_Cuonter, 4) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 5) = Choose(n_Cuonter + 1, False, False, False, False, False, False, False, False)
    aElemento(n_Cuonter, 6) = Choose(n_Cuonter + 1, True, True, True, True, True, True, True, True)
    aElemento(n_Cuonter, 7) = Choose(n_Cuonter + 1, "", "", "", "", "", "", "", "")
    aElemento(n_Cuonter, 8) = Choose(n_Cuonter + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Cuonter, 9) = Choose(n_Cuonter + 1, 1, 0, 0, 0, 0, 0, 0, 1)
  Next n_Cuonter
  
  ReDim aElementos(1, 3)
  For n_Cuonter = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Cuonter, 0) = ""
      aElementos(n_Cuonter, 1) = n_BackColorMdf: aElementos(n_Cuonter, 2) = vbBlack
  Next n_Cuonter
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbcomprobantes, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbcomprobantes, "Comprobantes de Cuarta Categoria", 3
  ']
  
  ' Carga los datos en el formulario
  ShowScreen
  ' Selecciono la primera pestaña
  n_IndexTabs = 2
  tasRegister.Tabs(1).Selected = True
  
  lblTitle = "Trabajador(a) " & " " & Trim(txtNombres(0)) & " " & Trim(txtNombres(1)) & " " & Trim(txtNombres(2))
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  
  If FormVisible("fAbcExperienciaLaboral") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcExperienciaLaboral.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
  If FormVisible("fAbcEstudioRealizado") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcEstudioRealizado.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
  If FormVisible("fAbcContrato") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcContrato.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
  If FormVisible("fAbcFamiliar") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcFamiliar.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
  If FormVisible("fAbcRemunAnterior") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcRemunAnterior.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
 If FormVisible("fAbcTerceros") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcTerceros.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
 If FormVisible("fAbcSuspension") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcSuspension.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If
 If FormVisible("fAbcComprobantes") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcComprobantes.Caption, vbExclamation
    Cancel = True
    Exit Sub
  End If

End Sub
Private Sub imgFoto_DblClick()
  
  On Error GoTo CancelaDialogo
  fMenu.cdlDialogo.DialogTitle = "Seleccionar Imagen"
  fMenu.cdlDialogo.CancelError = True
  fMenu.cdlDialogo.Flags = cdlOFNHideReadOnly
  fMenu.cdlDialogo.DefaultExt = ".bmp"
  fMenu.cdlDialogo.Filter = "Imagen BMP (*.bmp)|*.bmp|Imagen JPEG(*.jpg)|*.jpg|Imagen GIF (*.gif)|*.gif|Todos los archivos(*.*)|*.*"
  fMenu.cdlDialogo.FilterIndex = 1
  fMenu.cdlDialogo.ShowOpen
  imgFoto.Picture = LoadPicture(fMenu.cdlDialogo.FileName)
  imgFoto.Tag = fMenu.cdlDialogo.FileName
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then
    MsgBox Error(Err.Number)
    Exit Sub
  End If
  On Error GoTo 0
  
End Sub
Private Sub imgFoto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Elimino la fotografia
  If Button = vbRightButton And Shift = s_Estado_Ina Then
    If MsgBox("Desea Eliminar Fotografía del " & lblTitle, vbQuestion + vbYesNo) = vbYes Then
      imgFoto.Picture = LoadPicture("")
      imgFoto.Tag = ""
    End If
  End If
End Sub
Private Sub mskFecBaja_GotFocus()
  gdl_Procedure.MarcaGet mskFecBaja
End Sub
Private Sub mskFecBaja_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecBaja_LostFocus()
  If Left(cmbSituacion.Text, 1) = "I" Then mskFecSitua = Format(mskFecBaja, s_FormatoFecha)
End Sub
Private Sub mskFecBaja_Validate(Cancel As Boolean)
  If mskFecBaja.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecBaja, 1900
  End If
End Sub
Private Sub mskFecInicio_GotFocus()
  gdl_Procedure.MarcaGet mskFecInicio
End Sub
Private Sub mskFecInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecInicio_Validate(Cancel As Boolean)
  If mskFecInicio.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecInicio, 1900
  End If
End Sub
Private Sub mskFecSitua_GotFocus()
  gdl_Procedure.MarcaGet mskFecSitua
End Sub
Private Sub mskFecSitua_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecSitua_Validate(Cancel As Boolean)
  If mskFecSitua.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecSitua, 1900
  End If
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub tasRegister_Click()
  
  ' Valido selección de remuneración
  If (ps_NivelUsr = NivelUsuario.Auxiliar And tasRegister.SelectedItem.Index = 4) Then Beep: MsgBox "Información  Restringida " & lblTitle.Caption, vbInformation: tasRegister.Tabs(1).Selected = True: Exit Sub
  frmOpciones(n_IndexTabs).Left = -20000
  frmOpciones(n_IndexTabs).Enabled = False
  n_IndexTabs = (tasRegister.SelectedItem.Index - 1)
  frmOpciones(n_IndexTabs).Top = 525
  frmOpciones(n_IndexTabs).Left = 150
  frmOpciones(n_IndexTabs).Enabled = Not (ps_NivelUsr = NivelUsuario.Auxiliar And (n_IndexTabs = 0 Or n_IndexTabs = 2 Or n_IndexTabs = 3 Or n_IndexTabs = 8))

End Sub

Private Sub tdbAnterior_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionAnt_Click 2
End Sub

Private Sub tdbAnterior_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbAnterior.Bookmark Then Exit Sub
  If FormVisible("fAbcRemunAnterior") Then
    If Not tdbAnterior.EOF And Not tdbAnterior.BOF Then
      fAbcRemunAnterior.ShowScreen
    End If
  End If

End Sub

Private Sub tdbContrato_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionCon_Click 2
End Sub
Private Sub tdbContrato_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

  If LastRow = tdbContrato.Bookmark Then Exit Sub
  If FormVisible("fAbcContrato") Then
    If Not tdbContrato.EOF And Not tdbContrato.BOF Then
      fAbcContrato.ShowScreen
    End If
  End If

End Sub
Private Sub tdbEstudio_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionEst_Click 2
End Sub
Private Sub tdbEsta_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionEsta_Click 2
End Sub
Private Sub tdbEmpleador_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionEmp_Click 2
End Sub
Private Sub tdbEstudio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbEstudio.Bookmark Then Exit Sub
  If FormVisible("fAbcEstudioRealizado") Then
    If Not tdbEstudio.EOF And Not tdbEstudio.BOF Then fAbcEstudioRealizado.ShowScreen
  End If

End Sub
Private Sub tdbEsta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbesta.Bookmark Then Exit Sub
  If FormVisible("fAbcEstaLaboral") Then
    If Not tdbesta.EOF And Not tdbesta.BOF Then fAbcEstaLaboral.ShowScreen
  End If

End Sub
Private Sub tdbEmpleador_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbempleador.Bookmark Then Exit Sub
  If FormVisible("fAbcEmpleadores") Then
    If Not tdbempleador.EOF And Not tdbempleador.BOF Then fAbcEmpleadores.ShowScreen
  End If

End Sub
Private Sub tdbExperiencia_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionExp_Click 2
End Sub
Private Sub tdbExperiencia_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbExperiencia.Bookmark Then Exit Sub
  If FormVisible("fAbcExperienciaLaboral") Then
    If Not tdbExperiencia.EOF And Not tdbExperiencia.BOF Then fAbcExperienciaLaboral.ShowScreen
  End If

End Sub
Private Sub tdbFamiliar_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionFam_Click 2
End Sub
Private Sub tdbFamiliar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbFamiliar.Bookmark Then Exit Sub
  If FormVisible("fAbcFamiliar") Then
    If Not tdbFamiliar.EOF And Not tdbFamiliar.BOF Then fAbcFamiliar.ShowScreen
  End If

End Sub
Private Sub tdbTerceros_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionterceros_Click 2
End Sub
Private Sub tdbTerceros_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbterceros.Bookmark Then Exit Sub
  If FormVisible("fAbcTerceros") Then
    If Not tdbterceros.EOF And Not tdbterceros.BOF Then fAbcTerceros.ShowScreen
  End If

End Sub
Private Sub tdbSuspension_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActionsuspension_Click 2
End Sub
Private Sub tdbSuspension_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbsuspension.Bookmark Then Exit Sub
  If FormVisible("fAbcSuspension") Then
    If Not tdbsuspension.EOF And Not tdbsuspension.BOF Then fAbcSuspension.ShowScreen
  End If

End Sub
Private Sub tdbComprobantes_DblClick()
  If Me.Tag = s_MdoData_Upd Then cmdActioncomprobantes_Click 2
End Sub
Private Sub tdbComprobantes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  If LastRow = tdbcomprobantes.Bookmark Then Exit Sub
  If FormVisible("fAbccomprobantes") Then
    If Not tdbcomprobantes.EOF And Not tdbcomprobantes.BOF Then fAbcComprobantes.ShowScreen
  End If

End Sub
Private Sub tdbHelp_DblClick()
  
  If dcaHelp.Recordset.RecordCount = 0 Or (dcaHelp.Recordset.EOF And dcaHelp.Recordset.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Tipo de documento de identidad
    txtTipoDocu = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoDocu.SetFocus
   Case 1       ' Tipo de via
    txtTipoVia = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoVia.SetFocus
   Case 2       ' Tipo de zona
    txtTipoZona = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoZona.SetFocus
   Case 3       ' Tipo de trabajador
    txtTipoTraba = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoTraba.SetFocus
   Case 4       ' Cargo de trabajador
    txtCargo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCargo.SetFocus
   Case 5       ' Profesion
    txtProfesion = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtProfesion.SetFocus
   Case 6       ' Centro de costo
    txtCenCosto = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCenCosto.SetFocus
   Case 7       ' Entidad EPS
    txtCodEps = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCodEps.SetFocus
   Case 8       ' Entidad de pension
    txtEntidadAfp = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtEntidadAfp.SetFocus
   Case 9, 10, 24, 25  ' Entidad bancaria
    txtBanco(IIf(n_IndexHelp = 9, 0, IIf(n_IndexHelp = 24, 1, IIf(n_IndexHelp = 10, 2, 3)))).Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtBanco(IIf(n_IndexHelp = 9, 0, IIf(n_IndexHelp = 24, 1, IIf(n_IndexHelp = 10, 2, 3)))).SetFocus
   Case 11       ' Ubicación o localidad
    txtUbicacion = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtUbicacion.SetFocus
   Case 12       ' Sección de la empresa
    txtSeccion = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtSeccion.SetFocus
   Case 13, 14 ' Concepto de planilla remuneración
    txtConcepto(n_IndexHelp - 13) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtConcepto(n_IndexHelp - 13).SetFocus
    Case 15 ' Nacionalidad
    txtNacional = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtNacional.SetFocus
    Case 16 ' Situacion EPS
    txtSiteps = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtSiteps.SetFocus
    Case 17 ' Tipo de Pago
    txtTipPag = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipPag.SetFocus
    Case 18 ' Fin de Periodo
    txtfinperiodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtfinperiodo.SetFocus
    Case 19 ' MOdalidad Formativa
    txtmodformativa = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtmodformativa.SetFocus
    Case 20 ' Periodicidad
    txtPeriodicidad = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodicidad.SetFocus
    Case 21 ' Categoria Ocupacional
    txtcatocupacional = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtcatocupacional.SetFocus
    Case 22 ' Convenios por Tributacion
    txttributacion = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txttributacion.SetFocus
   Case 23       ' Condición d etrabajo
    txtCondicion.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCondicion.SetFocus
   Case 26       ' Larga distancia
    txtLargaDistancia.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtLargaDistancia.SetFocus
  End Select
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0  ' Tipo de documento de identidad
    s_Sql = gdl_Funcion.HelpTablas("dci", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 1  ' Tipo de via
    s_Sql = gdl_Funcion.HelpTablas("via", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 2  ' Tipo de zona
    s_Sql = gdl_Funcion.HelpTablas("zon", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 3     ' Tipo de trabajador
    s_Sql = gdl_Funcion.HelpTablas("tpt", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 4     ' Cargo de personal
    s_Sql = gdl_Funcion.HelpTablas("cgo", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
   Case 5     ' Profesión de personal
    s_Sql = gdl_Funcion.HelpTablas("pfs", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 6     ' Centro de costo
    s_Sql = gdl_Funcion.HelpTablas("cco", tdbHelp.Columns(ColIndex).DataField, pn_NivelCenCosto, "")
   Case 7     ' Entidad de EPS
    s_Sql = gdl_Funcion.HelpTablas("eps", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 8     ' Entidad de pensión
    s_Sql = gdl_Funcion.HelpTablas("afp", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 9, 10, 24   ' Entidad bancaria
    s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 11     ' Ubicación o localidad
    s_Sql = gdl_Funcion.HelpTablas("ubi", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 12     ' Sección de la empresa
    s_Sql = gdl_Funcion.HelpTablas("sec", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 13, 14  ' Conceptos de planilla remuneración
    s_Personal = ps_ClsPlanilla & IIf(n_IndexHelp = 10, "F", "C") & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Personal, "")
   Case 15 ' Nacionalidad
    s_Sql = gdl_Funcion.HelpTablas("nac", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 23     ' Condición de trabajo
    s_Sql = gdl_Funcion.HelpTablas("cdt", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
   Case 26     ' Larga distancia
    s_Sql = gdl_Funcion.HelpTablas("ldn", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
  End Select
  dcaHelp.RecordSource = s_Sql
  dcaHelp.Refresh

End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaHelp, tdbHelp, ""
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim porstClone As ADODB.Recordset, v_Bookmark As Variant
  Dim n_Columna As Integer, s_Criterio As String

  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    ' Conformo la cadena de ayuda
    s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    Set porstClone = dcaHelp.Recordset.Clone()
    
    n_Columna = tdbHelp.Col
    s_Criterio = tdbHelp.Columns(n_Columna).DataField & " >= '" & s_SqlHelp & "'"
    porstClone.Find s_Criterio, 0, adSearchForward, 0
    If Not (porstClone.BOF Or porstClone.EOF) Then
      dcaHelp.Recordset.Bookmark = porstClone.Bookmark
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
Private Sub tdbRegistro_AfterColUpdate(ByVal ColIndex As Integer)
  
  If ColIndex = 4 Then
    n_ImporteCol = CDec(lblTotalRemunera) - n_ImporteCol
    n_ImporteCol = n_ImporteCol + CDec(tdbRegistro.Columns(ColIndex).Value)
    lblTotalRemunera = FormatNumber(n_ImporteCol, 2)
  End If

End Sub
Private Sub tdbRegistro_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  
  If ColIndex = 4 Then
    If Not IsNumeric(tdbRegistro.Columns(ColIndex).Text) Then
      Beep
      MsgBox "Debe ingresar un Valor Numérico", vbExclamation
      Cancel = True
    End If
    n_ImporteCol = CDec(OldValue)
  End If

End Sub
Private Sub txtBanco_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtBanco(Index)
End Sub
Private Sub txtBanco_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click (Choose(Index + 1, 9, 24, 10, 25))
End Sub
Private Sub txtBanco_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtBanco_LostFocus(Index As Integer)
  lblHelp(Choose(Index + 1, 9, 24, 10, 25)) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(Index).Text, "EB")
End Sub

Private Sub txtCargo_GotFocus()
  gdl_Procedure.MarcaGet txtCargo
End Sub
Private Sub txtCargo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 4
End Sub
Private Sub txtCargo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCargo_LostFocus()
  lblHelp(4) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtCargo, "DC")
End Sub

Private Sub txtCenCosto_GotFocus()
  gdl_Procedure.MarcaGet txtCenCosto
End Sub
Private Sub txtCenCosto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 6
End Sub
Private Sub txtCenCosto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCenCosto_LostFocus()
  lblHelp(6) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, ps_CodEmpresa, txtCenCosto, "CC")
End Sub
Private Sub txtCodEps_GotFocus()
  gdl_Procedure.MarcaGet txtCodEps
End Sub
Private Sub txtCodEps_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 7
End Sub
Private Sub txtCodEps_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCodEps_LostFocus()
  lblHelp(7) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtCodEps, "ES")
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
      txtTipoDocu.SetFocus
      KeyAscii = 0
    End If
  End If
End Sub
Private Sub txtConcepto_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtConcepto(Index)
End Sub
Private Sub txtConcepto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click (13 + Index)
End Sub
Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtConcepto_LostFocus(Index As Integer)
  lblHelp(13 + Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtConcepto(Index), "CP")
End Sub
Private Sub txtCondicion_GotFocus()
  gdl_Procedure.MarcaGet txtCondicion
End Sub
Private Sub txtCondicion_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click (23)
End Sub
Private Sub txtCondicion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCondicion_LostFocus()
  lblHelp(23).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtCondicion.Text, "ST")
End Sub

Private Sub txtCuenta_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCuenta(Index)
End Sub
Private Sub txtCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtDependientes_GotFocus()
  gdl_Procedure.MarcaGet txtDependientes
End Sub
Private Sub txtDependientes_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtDependientes_Validate(Cancel As Boolean)
  txtDependientes.Text = IIf(Not IsNumeric(txtDependientes.Text), 0, txtDependientes.Text)
  txtDependientes.Text = IIf(CInt(txtDependientes.Text) < 0, 0, txtDependientes.Text)
  txtDependientes.Text = CInt(txtDependientes.Text)
End Sub
Private Sub txtDocumento_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtDocumento(Index)
End Sub
Private Sub txtDocumento_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub txtDocumento_LostFocus(Index As Integer)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim mensaje As String
If Index = 0 Then
    sql = "select codpsn,numdociden,concat(apepaterno,' ',apematerno,' ',nombres) as nombres,fecingreso,fecbaja from plpersonal where numdociden='" & txtDocumento(0) & "' and codcls='" & ps_ClsPlanilla & "'"
    rs.Open sql, cnn, adOpenDynamic, adLockPessimistic
    Do Until rs.EOF
        mensaje = mensaje & " Codigo : (" & rs(0).Value & ") " & rs(2).Value & "  Ingreso el " & rs(3).Value & " Ceso el " & IIf(IsNull(rs(4).Value) = True, Space(10), rs(4).Value) & vbCrLf
        rs.MoveNext
    Loop
    If mensaje <> "" Then
        MsgBox "Existen Registros con este Documento : " & vbCrLf & vbCrLf & mensaje, vbCritical, " Alerta"
        Beep
    End If
    rs.Close
End If
End Sub

Private Sub txtEntidadAfp_GotFocus()
  gdl_Procedure.MarcaGet txtEntidadAfp
End Sub
Private Sub txtEntidadAfp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 8
End Sub
Private Sub txtEntidadAfp_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtEntidadAfp_LostFocus()
  lblHelp(8) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtEntidadAfp, "EP")
End Sub
Private Sub txtEssalud_GotFocus()
  gdl_Procedure.MarcaGet txtEssalud
End Sub
Private Sub txtEssalud_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtHijos_GotFocus()
  gdl_Procedure.MarcaGet txtHijos
End Sub
Private Sub txtHijos_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtDependientes.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtHijos_Validate(Cancel As Boolean)
  txtHijos.Text = IIf(Not IsNumeric(txtHijos.Text), 0, txtHijos.Text)
  txtHijos.Text = IIf(CInt(txtHijos.Text) < 0, 0, txtHijos.Text)
  txtHijos.Text = CInt(txtHijos.Text)
End Sub
Private Sub txtJornadaLabor_GotFocus()
  gdl_Procedure.MarcaGet txtJornadaLabor
End Sub
Private Sub txtJornadaLabor_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtJornadaLabor_Validate(Cancel As Boolean)
  txtJornadaLabor.Text = IIf(Not IsNumeric(txtJornadaLabor.Text), 0, txtJornadaLabor.Text)
  txtJornadaLabor.Text = IIf(CDec(txtJornadaLabor.Text) < 0, 0, txtJornadaLabor.Text)
  txtJornadaLabor.Text = FormatNumber(txtJornadaLabor.Text, 2)
End Sub
Private Sub txtLargaDistancia_GotFocus()
  gdl_Procedure.MarcaGet txtLargaDistancia
End Sub
Private Sub txtLargaDistancia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtLargaDistancia_LostFocus()
lblHelp(26).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtLargaDistancia.Text, "LD")
End Sub
Private Sub txtNacional_GotFocus()
  gdl_Procedure.MarcaGet txtNacional
End Sub
Private Sub txtNacional_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtNacional_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNacional_LostFocus()
lblHelp(15) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtNacional, "NA")
End Sub
Private Sub txtNombres_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtNombres(Index)
End Sub
Private Sub txtNombres_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNombres_LostFocus(Index As Integer)
  For n_Cuonter = 0 To 5
    lblNombre(n_Cuonter) = "Nombres : " & Trim(txtNombres(0)) & " " & Trim(txtNombres(1)) & "; " & Trim(txtNombres(2)) & "  "
  Next n_Cuonter
End Sub
Private Sub txtNombreVia_GotFocus()
  gdl_Procedure.MarcaGet txtNombreVia
End Sub
Private Sub txtNombreVia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtnumero(0).SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtNombreZona_GotFocus()
  gdl_Procedure.MarcaGet txtNombreZona
End Sub
Private Sub txtNombreZona_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtReferencia.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtNroCuenta_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtNroCuenta(Index)
End Sub
Private Sub txtNroCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtnumero_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtnumero(Index)
End Sub
Private Sub txtnumero_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNumeroAfp_GotFocus()
  gdl_Procedure.MarcaGet txtNumeroAfp
End Sub
Private Sub txtNumeroAfp_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPorJudicial_GotFocus()
  gdl_Procedure.MarcaGet txtPorJudicial
End Sub
Private Sub txtPorJudicial_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPorJudicial_Validate(Cancel As Boolean)
  txtPorJudicial.Text = IIf(Not IsNumeric(txtPorJudicial.Text), 0, txtPorJudicial.Text)
  txtPorJudicial.Text = IIf(CDec(txtPorJudicial.Text) < 0, 0, txtPorJudicial.Text)
  txtPorJudicial.Text = FormatNumber(txtPorJudicial.Text, 2)
End Sub
Private Sub txtProfesion_GotFocus()
  gdl_Procedure.MarcaGet txtProfesion
End Sub
Private Sub txtProfesion_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 5
End Sub
Private Sub txtProfesion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtProfesion_LostFocus()
  lblHelp(5) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtProfesion, "PF")
End Sub
Private Sub txtReferencia_GotFocus()
  gdl_Procedure.MarcaGet txtReferencia
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtTelefono(0).SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemuNeta_GotFocus()
  gdl_Procedure.MarcaGet txtRemuNeta
End Sub
Private Sub txtRemuNeta_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemuNeta_Validate(Cancel As Boolean)
  txtRemuNeta.Text = IIf(Not IsNumeric(txtRemuNeta.Text), 0, txtRemuNeta.Text)
  txtRemuNeta.Text = IIf(CDec(txtRemuNeta.Text) < 0, 0, txtRemuNeta.Text)
  txtRemuNeta.Text = FormatNumber(txtRemuNeta.Text, 2)
End Sub
Private Sub txtSeccion_GotFocus()
  gdl_Procedure.MarcaGet txtSeccion
End Sub
Private Sub txtSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 12
End Sub
Private Sub txtSeccion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtSeccion_LostFocus()
  lblHelp(12) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtSeccion, "SE")
End Sub
Private Sub txtTelefono_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtTelefono(Index)
End Sub
Private Sub txtTelefono_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoDocu_GotFocus()
  gdl_Procedure.MarcaGet txtTipoDocu
End Sub
Private Sub txtTipoDocu_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtTipoDocu_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtDocumento(0).SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoDocu_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoDocu, "DI")
End Sub
Private Sub txtTipoTraba_GotFocus()
  gdl_Procedure.MarcaGet txtTipoTraba
End Sub
Private Sub txtTipoTraba_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 3
End Sub
Private Sub txtTipoTraba_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoTraba_LostFocus()
  lblHelp(3) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoTraba, "TT")
End Sub
Private Sub txtTipoVia_GotFocus()
  gdl_Procedure.MarcaGet txtTipoVia
End Sub
Private Sub txtTipoVia_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtTipoVia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtNombreVia.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoVia_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoVia, "TV")
End Sub
Private Sub txtTipoZona_GotFocus()
  gdl_Procedure.MarcaGet txtTipoZona
End Sub
Private Sub txtTipoZona_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 2
End Sub
Private Sub txtTipoZona_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtNombreZona.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoZona_LostFocus()
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoZona, "TZ")
End Sub

Private Sub txtUbicacion_GotFocus()
  gdl_Procedure.MarcaGet txtUbicacion
End Sub
Private Sub txtUbicacion_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdUbigeo_Click 11
End Sub
Private Sub txtUbicacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtUbicacion_LostFocus()
  lblHelp(11) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtUbicacion, "UL")
End Sub
Private Sub txtUbigeo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtUbigeo(Index)
End Sub
Private Sub txtUbigeo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdUbigeo_Click 0
End Sub
Private Sub txtUbigeo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtUbigeo_LostFocus(Index As Integer)
  lblUbigeo(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUbigeo(Index), "UG")
End Sub
Private Sub txtfinperiodo_GotFocus()
  gdl_Procedure.MarcaGet txtfinperiodo
End Sub
Private Sub txtfinperiodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 5
End Sub
Private Sub txtfinperiodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtfinperiodo_LostFocus()
  lblHelp(18) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtfinperiodo, "FP")
End Sub
Private Sub txtmodformativa_GotFocus()
  gdl_Procedure.MarcaGet txtmodformativa
End Sub
Private Sub txtmodformativa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 5
End Sub
Private Sub txtmodformativa_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtmodformativa_LostFocus()
  lblHelp(19) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtmodformativa, "MF")
End Sub
Private Sub txtperiodicidad_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodicidad
End Sub
Private Sub txtperiodicidad_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 5
End Sub
Private Sub txtperiodicidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtperiodicidad_LostFocus()
  lblHelp(20) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtPeriodicidad, "PE")
End Sub
Private Sub txttippag_GotFocus()
  gdl_Procedure.MarcaGet txtTipPag
End Sub
Private Sub txttippag_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 5
End Sub
Private Sub txttippag_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txttippag_LostFocus()
  lblHelp(17) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipPag, "PA")
End Sub
Private Sub txtsiteps_GotFocus()
  gdl_Procedure.MarcaGet txtSiteps
End Sub
Private Sub txtsiteps_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 5
End Sub
Private Sub txtsiteps_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtsiteps_LostFocus()
  lblHelp(16) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtSiteps, "SI")
End Sub
Private Sub txtcatocupacional_GotFocus()
  gdl_Procedure.MarcaGet txtcatocupacional
End Sub
Private Sub txtcatocupacional_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 21
End Sub
Private Sub txtcatocupacional_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtcatocupacional_LostFocus()
  lblHelp(21) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtcatocupacional, "CY")
End Sub
Private Sub txttributacion_GotFocus()
  gdl_Procedure.MarcaGet txtcatocupacional
End Sub
Private Sub txttributacion_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 22
End Sub
Private Sub txttributacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txttributacion_LostFocus()
  lblHelp(22) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txttributacion, "CZ")
End Sub

