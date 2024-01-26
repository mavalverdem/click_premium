VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame frmRegister 
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   9657
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
         Height          =   4860
         Index           =   8
         Left            =   -80000
         TabIndex        =   1
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
            TabIndex        =   2
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
            TabIndex        =   3
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
               TabIndex        =   4
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
               TabIndex        =   5
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
               Picture         =   "prueba.frx":0000
            End
            Begin Threed.SSCommand cmdActionAnt 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   6
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
               Picture         =   "prueba.frx":001C
            End
            Begin Threed.SSCommand cmdActionAnt 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   7
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
               Picture         =   "prueba.frx":0038
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
            TabIndex        =   8
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   7
         Left            =   -70000
         TabIndex        =   9
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
            TabIndex        =   10
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
            TabIndex        =   11
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
               TabIndex        =   12
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
               TabIndex        =   13
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
               Picture         =   "prueba.frx":0054
            End
            Begin Threed.SSCommand cmdActionFam 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   14
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
               Picture         =   "prueba.frx":0070
            End
            Begin Threed.SSCommand cmdActionFam 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   15
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
               Picture         =   "prueba.frx":008C
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
            TabIndex        =   16
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   5
         Left            =   -50000
         TabIndex        =   17
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
            TabIndex        =   18
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
            TabIndex        =   19
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
               TabIndex        =   20
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
               TabIndex        =   21
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
               Picture         =   "prueba.frx":00A8
            End
            Begin Threed.SSCommand cmdActionEst 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   22
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
               Picture         =   "prueba.frx":00C4
            End
            Begin Threed.SSCommand cmdActionEst 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   23
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
               Picture         =   "prueba.frx":00E0
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
            TabIndex        =   24
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   4
         Left            =   -40000
         TabIndex        =   25
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
            TabIndex        =   26
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
            TabIndex        =   27
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
               TabIndex        =   28
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
               TabIndex        =   29
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
               Picture         =   "prueba.frx":00FC
            End
            Begin Threed.SSCommand cmdActionExp 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   30
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
               Picture         =   "prueba.frx":0118
            End
            Begin Threed.SSCommand cmdActionExp 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   31
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
               Picture         =   "prueba.frx":0134
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
            TabIndex        =   32
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   6
         Left            =   -60000
         TabIndex        =   33
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
            TabIndex        =   34
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
            TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
               Picture         =   "prueba.frx":0150
            End
            Begin Threed.SSCommand cmdActionCon 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   38
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
               Picture         =   "prueba.frx":016C
            End
            Begin Threed.SSCommand cmdActionCon 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   39
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
               Picture         =   "prueba.frx":0188
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
            TabIndex        =   40
            Top             =   465
            Width           =   840
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   3
         Left            =   -30000
         TabIndex        =   41
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   44
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
            Begin VB.TextBox txtConcepto 
               Height          =   280
               Index           =   0
               Left            =   975
               TabIndex        =   47
               Top             =   300
               Width           =   555
            End
            Begin VB.TextBox txtConcepto 
               Height          =   280
               Index           =   1
               Left            =   975
               TabIndex        =   46
               Top             =   990
               Width           =   555
            End
            Begin VB.TextBox txtRemuNeta 
               Height          =   280
               Left            =   975
               TabIndex        =   45
               Top             =   645
               Width           =   1185
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   13
               Left            =   1605
               TabIndex        =   48
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
               TabIndex        =   49
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
               Caption         =   "Neto :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   33
               Left            =   150
               TabIndex        =   54
               Top             =   300
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
               TabIndex        =   53
               Top             =   330
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Reajuste :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   35
               Left            =   150
               TabIndex        =   52
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
               Index           =   14
               Left            =   1935
               TabIndex        =   51
               Top             =   1020
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Importe :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   39
               Left            =   150
               TabIndex        =   50
               Top             =   645
               Width           =   735
            End
         End
         Begin Threed.SSCheck chkRemIntegral 
            Height          =   195
            Index           =   2
            Left            =   4395
            TabIndex        =   55
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
            TabIndex        =   56
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
            TabIndex        =   57
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
            TabIndex        =   58
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
            TabIndex        =   61
            Top             =   180
            Width           =   840
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
            TabIndex        =   60
            Top             =   3330
            Width           =   1245
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
            TabIndex        =   59
            Top             =   3375
            Width           =   1710
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
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   2
         Left            =   -20000
         TabIndex        =   62
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
         Begin VB.TextBox txtTipoTraba 
            Height          =   280
            Left            =   120
            MaxLength       =   2
            TabIndex        =   83
            Top             =   945
            Width           =   500
         End
         Begin VB.TextBox txtProfesion 
            Height          =   280
            Left            =   120
            MaxLength       =   2
            TabIndex        =   82
            Top             =   2370
            Width           =   500
         End
         Begin VB.TextBox txtCargo 
            Height          =   280
            Left            =   90
            MaxLength       =   2
            TabIndex        =   81
            Top             =   1485
            Width           =   500
         End
         Begin VB.TextBox txtCenCosto 
            Height          =   280
            Left            =   90
            MaxLength       =   20
            TabIndex        =   80
            Top             =   2955
            Width           =   900
         End
         Begin VB.ComboBox cmbSituacion 
            Height          =   315
            ItemData        =   "prueba.frx":01A4
            Left            =   6195
            List            =   "prueba.frx":01A6
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   3495
            Width           =   1710
         End
         Begin VB.TextBox txtEntidadAfp 
            Height          =   280
            Left            =   3225
            MaxLength       =   2
            TabIndex        =   78
            Top             =   390
            Width           =   500
         End
         Begin VB.TextBox txtNumeroAfp 
            Height          =   280
            Left            =   6360
            TabIndex        =   77
            Top             =   390
            Width           =   1530
         End
         Begin VB.TextBox txtCodEps 
            Height          =   280
            Left            =   90
            MaxLength       =   20
            TabIndex        =   66
            Top             =   3555
            Width           =   500
         End
         Begin VB.ComboBox cmbCargoConfi 
            Height          =   315
            ItemData        =   "prueba.frx":01A8
            Left            =   1950
            List            =   "prueba.frx":01AA
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1845
            Width           =   1260
         End
         Begin VB.TextBox txtUbicacion 
            Height          =   280
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   64
            Top             =   2925
            Width           =   500
         End
         Begin VB.TextBox txtSeccion 
            Height          =   280
            Left            =   5505
            MaxLength       =   2
            TabIndex        =   63
            Top             =   2925
            Width           =   500
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   870
            Index           =   4
            Left            =   90
            TabIndex        =   67
            Top             =   3870
            Width           =   3180
            _Version        =   65536
            _ExtentX        =   5609
            _ExtentY        =   1535
            _StockProps     =   14
            Caption         =   " Otros Datos "
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
            Begin Threed.SSCheck chkEssaludVida 
               Height          =   225
               Left            =   120
               TabIndex        =   68
               Top             =   225
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   397
               _StockProps     =   78
               Caption         =   "EsSalud Vida"
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
            Begin Threed.SSCheck chkSctr 
               Height          =   225
               Left            =   1605
               TabIndex        =   69
               Top             =   240
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2558
               _ExtentY        =   397
               _StockProps     =   78
               Caption         =   "Cobertura SCTR"
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
               Left            =   120
               TabIndex        =   70
               Top             =   510
               Width           =   1650
               _Version        =   65536
               _ExtentX        =   2910
               _ExtentY        =   397
               _StockProps     =   78
               Caption         =   "Afiliado a Sindicato"
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
         Begin Threed.SSFrame frmCuadro 
            Height          =   1425
            Index           =   5
            Left            =   3345
            TabIndex        =   71
            Top             =   3315
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
               TabIndex        =   72
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
               TabIndex        =   73
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
               TabIndex        =   74
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
               TabIndex        =   75
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
               TabIndex        =   76
               Top             =   1050
               Width           =   1125
            End
         End
         Begin MSComCtl2.DTPicker dtpFecIngreso 
            Height          =   285
            Left            =   120
            TabIndex        =   84
            Top             =   390
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Format          =   64487425
            CurrentDate     =   37515
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   3
            Left            =   675
            TabIndex        =   85
            Top             =   945
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCheck chkCargoConfi 
            Height          =   225
            Left            =   120
            TabIndex        =   86
            Top             =   1845
            Width           =   1740
            _Version        =   65536
            _ExtentX        =   3069
            _ExtentY        =   397
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
            Left            =   675
            TabIndex        =   87
            Top             =   2370
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
            Left            =   645
            TabIndex        =   88
            Top             =   1485
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
            Left            =   1065
            TabIndex        =   89
            Top             =   2955
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   1920
            Index           =   2
            Left            =   3225
            TabIndex        =   90
            Top             =   705
            Width           =   4695
            _Version        =   65536
            _ExtentX        =   8281
            _ExtentY        =   3387
            _StockProps     =   14
            Caption         =   " Detalle de Pago "
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
            Begin VB.TextBox txtNroCuenta 
               Height          =   280
               Index           =   1
               Left            =   2595
               MaxLength       =   20
               TabIndex        =   94
               Top             =   1545
               Width           =   1980
            End
            Begin VB.TextBox txtBanco 
               Height          =   280
               Index           =   1
               Left            =   180
               MaxLength       =   2
               TabIndex        =   93
               Top             =   1545
               Width           =   500
            End
            Begin VB.TextBox txtNroCuenta 
               Height          =   280
               Index           =   0
               Left            =   2595
               MaxLength       =   20
               TabIndex        =   92
               Top             =   975
               Width           =   1980
            End
            Begin VB.TextBox txtBanco 
               Height          =   280
               Index           =   0
               Left            =   180
               MaxLength       =   2
               TabIndex        =   91
               Top             =   975
               Width           =   500
            End
            Begin Threed.SSCheck chkPagoDolar 
               Height          =   285
               Left            =   180
               TabIndex        =   95
               Top             =   390
               Width           =   1905
               _Version        =   65536
               _ExtentX        =   3351
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "Pago en Dólares"
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
               Index           =   9
               Left            =   735
               TabIndex        =   96
               Top             =   975
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
               Index           =   10
               Left            =   735
               TabIndex        =   97
               Top             =   1545
               Width           =   280
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin Threed.SSCheck chkCtsDolar 
               Height          =   285
               Left            =   2595
               TabIndex        =   98
               Top             =   390
               Width           =   1960
               _Version        =   65536
               _ExtentX        =   3466
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
            Begin Threed.SSCheck chkCtsDeposito 
               Height          =   285
               Left            =   2595
               TabIndex        =   99
               Top             =   120
               Width           =   1960
               _Version        =   65536
               _ExtentX        =   3457
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
            Begin VB.Label lblDato 
               Caption         =   "Número de Cuenta :"
               ForeColor       =   &H00000000&
               Height          =   190
               Index           =   26
               Left            =   2595
               TabIndex        =   105
               Top             =   1305
               Width           =   1965
            End
            Begin VB.Label lblDato 
               Caption         =   "Banco de C.T.S. :"
               ForeColor       =   &H00000000&
               Height          =   190
               Index           =   25
               Left            =   180
               TabIndex        =   104
               Top             =   1305
               Width           =   1545
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
               Left            =   1065
               TabIndex        =   103
               Top             =   1590
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Número de Cuenta :"
               ForeColor       =   &H00000000&
               Height          =   190
               Index           =   24
               Left            =   2595
               TabIndex        =   102
               Top             =   735
               Width           =   1965
            End
            Begin VB.Label lblDato 
               Caption         =   "Banco de Pago :"
               ForeColor       =   &H00000000&
               Height          =   190
               Index           =   27
               Left            =   180
               TabIndex        =   101
               Top             =   735
               Width           =   1545
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
               Left            =   1065
               TabIndex        =   100
               Top             =   1020
               Width           =   195
            End
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   8
            Left            =   3780
            TabIndex        =   106
            Top             =   390
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin MSMask.MaskEdBox mskFecBaja 
            Height          =   285
            Left            =   6195
            TabIndex        =   107
            Top             =   4455
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
            Left            =   645
            TabIndex        =   108
            Top             =   3555
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
            Left            =   3795
            TabIndex        =   109
            Top             =   2925
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
            Left            =   6060
            TabIndex        =   110
            Top             =   2925
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
            Left            =   6840
            TabIndex        =   111
            Top             =   3885
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
            TabIndex        =   133
            Top             =   450
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Ingreso :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   120
            TabIndex        =   132
            Top             =   150
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
            Index           =   3
            Left            =   1020
            TabIndex        =   131
            Top             =   990
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Trabajador :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   130
            Top             =   705
            Width           =   1470
         End
         Begin VB.Label lblDato 
            Caption         =   "Cargo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   120
            TabIndex        =   129
            Top             =   1245
            Width           =   1470
         End
         Begin VB.Label lblDato 
            Caption         =   "Profesión u Oficio :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   128
            Top             =   2115
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
            Index           =   5
            Left            =   1005
            TabIndex        =   127
            Top             =   2415
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
            Left            =   975
            TabIndex        =   126
            Top             =   1530
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Centro de Costo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   105
            TabIndex        =   125
            Top             =   2700
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
            Left            =   1395
            TabIndex        =   124
            Top             =   3000
            Width           =   210
         End
         Begin VB.Label lblDato 
            Caption         =   "Situación :"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   34
            Left            =   6195
            TabIndex        =   123
            Top             =   3285
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
            Index           =   8
            Left            =   4110
            TabIndex        =   122
            Top             =   420
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Entidad de Pensión :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   3225
            TabIndex        =   121
            Top             =   150
            Width           =   1545
         End
         Begin VB.Label lblDato 
            Caption         =   "Número de Pensión :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   6360
            TabIndex        =   120
            Top             =   150
            Width           =   1530
         End
         Begin VB.Label lblDato 
            Caption         =   "EPS :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   105
            TabIndex        =   119
            Top             =   3315
            Width           =   1470
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Baja :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   43
            Left            =   6195
            TabIndex        =   118
            Top             =   4230
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
            Left            =   975
            TabIndex        =   117
            Top             =   3585
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
            Index           =   11
            Left            =   4125
            TabIndex        =   116
            Top             =   2970
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Ubicación :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   3270
            TabIndex        =   115
            Top             =   2670
            Width           =   1470
         End
         Begin VB.Label lblDato 
            Caption         =   "Sección :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   5535
            TabIndex        =   114
            Top             =   2670
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
            Left            =   6390
            TabIndex        =   113
            Top             =   2970
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha :"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   48
            Left            =   6195
            TabIndex        =   112
            Top             =   3930
            Width           =   1335
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   1
         Left            =   -10000
         TabIndex        =   134
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
         Begin VB.TextBox txtUbigeo 
            Height          =   280
            Index           =   1
            Left            =   300
            TabIndex        =   142
            Top             =   3570
            Width           =   975
         End
         Begin VB.TextBox txtReferencia 
            Height          =   750
            Left            =   300
            MultiLine       =   -1  'True
            TabIndex        =   141
            Top             =   2490
            Width           =   3600
         End
         Begin VB.TextBox txtNombreZona 
            Height          =   280
            Left            =   4260
            TabIndex        =   140
            Top             =   1875
            Width           =   3470
         End
         Begin VB.TextBox txtTipoZona 
            Height          =   280
            Left            =   315
            TabIndex        =   139
            Top             =   1875
            Width           =   500
         End
         Begin VB.TextBox txtNumero 
            Height          =   280
            Index           =   1
            Left            =   1365
            TabIndex        =   138
            Top             =   1260
            Width           =   870
         End
         Begin VB.TextBox txtNumero 
            Height          =   280
            Index           =   0
            Left            =   315
            TabIndex        =   137
            Top             =   1260
            Width           =   870
         End
         Begin VB.TextBox txtNombreVia 
            Height          =   280
            Left            =   4260
            TabIndex        =   136
            Top             =   630
            Width           =   3470
         End
         Begin VB.TextBox txtTipoVia 
            Height          =   280
            Left            =   315
            TabIndex        =   135
            Top             =   630
            Width           =   500
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   870
            Index           =   1
            Left            =   4260
            TabIndex        =   143
            Top             =   2385
            Width           =   3270
            _Version        =   65536
            _ExtentX        =   5768
            _ExtentY        =   1535
            _StockProps     =   14
            Caption         =   " Numeros Teléfonicos "
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
            Begin VB.TextBox txtTelefono 
               Height          =   280
               Index           =   0
               Left            =   180
               TabIndex        =   145
               Top             =   495
               Width           =   1290
            End
            Begin VB.TextBox txtTelefono 
               Height          =   280
               Index           =   1
               Left            =   1800
               TabIndex        =   144
               Top             =   495
               Width           =   1290
            End
            Begin VB.Label lblDato 
               Caption         =   "Fijo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   16
               Left            =   180
               TabIndex        =   147
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Móvil :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   11
               Left            =   1800
               TabIndex        =   146
               Top             =   240
               Width           =   1335
            End
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   280
            Index           =   1
            Left            =   870
            TabIndex        =   148
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
            TabIndex        =   149
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
            Height          =   280
            Index           =   1
            Left            =   1335
            TabIndex        =   150
            Top             =   3570
            Width           =   280
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
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
            TabIndex        =   161
            Top             =   3615
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Ubigeo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   300
            TabIndex        =   160
            Top             =   3315
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Referencia :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   300
            TabIndex        =   159
            Top             =   2235
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombre de Zona :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4260
            TabIndex        =   158
            Top             =   1620
            Width           =   2280
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
            TabIndex        =   157
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Zona :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   315
            TabIndex        =   156
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
            Index           =   1
            Left            =   1275
            TabIndex        =   155
            Top             =   675
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Interior :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   1365
            TabIndex        =   154
            Top             =   1005
            Width           =   870
         End
         Begin VB.Label lblDato 
            Caption         =   "Número :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   315
            TabIndex        =   153
            Top             =   1005
            Width           =   870
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombre de Vía :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   4260
            TabIndex        =   152
            Top             =   375
            Width           =   2280
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Vía :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   315
            TabIndex        =   151
            Top             =   375
            Width           =   1335
         End
      End
      Begin Threed.SSFrame frmOpciones 
         Height          =   4860
         Index           =   0
         Left            =   -9000
         TabIndex        =   162
         Top             =   525
         Width           =   8025
         _Version        =   65536
         _ExtentX        =   14155
         _ExtentY        =   8572
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
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Index           =   1
            Left            =   6030
            TabIndex        =   172
            Top             =   825
            Width           =   1530
         End
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Index           =   0
            Left            =   1605
            TabIndex        =   171
            Top             =   825
            Width           =   1440
         End
         Begin VB.TextBox txtTipoDocu 
            Height          =   280
            Left            =   1605
            TabIndex        =   170
            Top             =   450
            Width           =   500
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
            TabIndex        =   169
            Top             =   465
            Width           =   1320
         End
         Begin VB.TextBox txtHijos 
            Height          =   280
            Left            =   120
            TabIndex        =   168
            Top             =   4440
            Width           =   800
         End
         Begin VB.TextBox txtDependientes 
            Height          =   280
            Left            =   1650
            TabIndex        =   167
            Top             =   4440
            Width           =   800
         End
         Begin VB.TextBox txtPorJudicial 
            Height          =   280
            Left            =   2925
            TabIndex        =   166
            Top             =   4440
            Width           =   800
         End
         Begin VB.TextBox txtEssalud 
            Height          =   280
            Left            =   6030
            TabIndex        =   165
            Top             =   3690
            Width           =   1785
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   0
            Left            =   4965
            TabIndex        =   164
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   1
            Left            =   6465
            TabIndex        =   163
            Top             =   4440
            Width           =   1200
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   2985
            Index           =   0
            Left            =   75
            TabIndex        =   173
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
            Begin VB.ComboBox cmbEstadoCivil 
               Height          =   315
               ItemData        =   "prueba.frx":01AC
               Left            =   3015
               List            =   "prueba.frx":01AE
               Style           =   2  'Dropdown List
               TabIndex        =   180
               Top             =   1665
               Width           =   1485
            End
            Begin VB.ComboBox cmbSexo 
               Height          =   315
               ItemData        =   "prueba.frx":01B0
               Left            =   3015
               List            =   "prueba.frx":01B2
               Style           =   2  'Dropdown List
               TabIndex        =   179
               Top             =   1065
               Width           =   1500
            End
            Begin VB.TextBox txtNombres 
               Height          =   280
               Index           =   2
               Left            =   150
               TabIndex        =   178
               Top             =   1065
               Width           =   2700
            End
            Begin VB.TextBox txtNombres 
               Height          =   280
               Index           =   0
               Left            =   150
               TabIndex        =   177
               Top             =   480
               Width           =   2700
            End
            Begin VB.TextBox txtNombres 
               Height          =   280
               Index           =   1
               Left            =   3015
               TabIndex        =   176
               Top             =   480
               Width           =   2700
            End
            Begin VB.TextBox txtUbigeo 
               Height          =   280
               Index           =   0
               Left            =   180
               TabIndex        =   175
               Top             =   2280
               Width           =   975
            End
            Begin VB.TextBox txtNacional 
               Height          =   280
               Left            =   1395
               TabIndex        =   174
               Top             =   2625
               Width           =   2700
            End
            Begin MSComCtl2.DTPicker dtpFecha 
               Height          =   285
               Left            =   150
               TabIndex        =   181
               Top             =   1665
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   503
               _Version        =   393216
               Format          =   64487425
               CurrentDate     =   37515
            End
            Begin Threed.SSCommand cmdUbigeo 
               Height          =   300
               Index           =   0
               Left            =   1215
               TabIndex        =   182
               Top             =   2280
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin VB.Label lblDato 
               Caption         =   "Estado Civil :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   8
               Left            =   3015
               TabIndex        =   192
               Top             =   1455
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Sexo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   3015
               TabIndex        =   191
               Top             =   855
               Width           =   1005
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
               TabIndex        =   190
               Top             =   1725
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Fecha de Nacimiento :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   150
               TabIndex        =   189
               Top             =   1455
               Width           =   1680
            End
            Begin VB.Label lblDato 
               Caption         =   "Apellido Paterno :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   3
               Left            =   150
               TabIndex        =   188
               Top             =   255
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Apellido Materno :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   4
               Left            =   3015
               TabIndex        =   187
               Top             =   255
               Width           =   1335
            End
            Begin VB.Label lblDato 
               Caption         =   "Nombres :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   150
               TabIndex        =   186
               Top             =   855
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
               Index           =   0
               Left            =   1620
               TabIndex        =   185
               Top             =   2340
               Width           =   195
            End
            Begin VB.Label lblDato 
               Caption         =   "Lugar de Nacimiento :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   18
               Left            =   180
               TabIndex        =   184
               Top             =   2055
               Width           =   1755
            End
            Begin VB.Label lblDato 
               Caption         =   "Nacionalidad :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   47
               Left            =   180
               TabIndex        =   183
               Top             =   2670
               Width           =   1080
            End
         End
         Begin Threed.SSCheck chkExtanjero 
            Height          =   300
            Left            =   6030
            TabIndex        =   193
            Top             =   465
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
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
            TabIndex        =   194
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
            TabIndex        =   195
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
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Left            =   6030
            Stretch         =   -1  'True
            ToolTipText     =   "Haga doble click para fotografía"
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label lblDato 
            Caption         =   "Documento Militar :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   4530
            TabIndex        =   205
            Top             =   870
            Width           =   1380
         End
         Begin VB.Label lblDato 
            Caption         =   "Documento de Identificación :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1605
            TabIndex        =   204
            Top             =   210
            Width           =   2070
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
            TabIndex        =   203
            Top             =   495
            Width           =   195
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
            TabIndex        =   202
            Top             =   210
            Width           =   960
         End
         Begin VB.Label lblDato 
            Caption         =   "Fotografía :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   6030
            TabIndex        =   201
            Top             =   1170
            Width           =   1680
         End
         Begin VB.Label lblDato 
            Caption         =   "Número de Hijos :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   200
            Top             =   4215
            Width           =   1305
         End
         Begin VB.Label lblDato 
            Caption         =   "Dependientes :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   1650
            TabIndex        =   199
            Top             =   4215
            Width           =   1110
         End
         Begin VB.Label lblDato 
            Caption         =   "Carnet Essalud :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   44
            Left            =   6030
            TabIndex        =   198
            Top             =   3450
            Width           =   1530
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Deudor :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   45
            Left            =   4980
            TabIndex        =   197
            Top             =   4215
            Width           =   1170
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Acreedor :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   6480
            TabIndex        =   196
            Top             =   4215
            Width           =   1290
         End
      End
      Begin MSComctlLib.TabStrip tasRegister 
         Height          =   5175
         Left            =   0
         TabIndex        =   206
         Top             =   120
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9128
         Style           =   2
         Separators      =   -1  'True
         TabMinWidth     =   1176
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   9
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
