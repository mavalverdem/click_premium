VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fPrmCertifik5ta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4860
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "prmcerti5ta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6420
   Begin TabDlg.SSTab tabRegister 
      Height          =   3705
      Left            =   75
      TabIndex        =   18
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   6535
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
      TabPicture(0)   =   "prmcerti5ta.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHelp(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblHelp(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblHelp(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblHelp(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDato(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDato(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblHelp(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "shpCuadro(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdHelp(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdHelp(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdHelp(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdHelp(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdHelp(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtRemunera(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtQuinta(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtUit"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtRemunera(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtQuinta(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      Begin VB.TextBox txtQuinta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   2220
         Width           =   980
      End
      Begin VB.TextBox txtRemunera 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   980
      End
      Begin VB.TextBox txtUit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   240
         TabIndex        =   13
         Top             =   2790
         Width           =   980
      End
      Begin VB.TextBox txtQuinta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1650
         Width           =   980
      End
      Begin VB.TextBox txtRemunera 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   510
         Width           =   980
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   1290
         TabIndex        =   19
         Top             =   510
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   2
         Left            =   1290
         TabIndex        =   21
         Top             =   1650
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   4
         Left            =   1290
         TabIndex        =   23
         Top             =   2790
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   1
         Left            =   1290
         TabIndex        =   20
         Top             =   1080
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   3
         Left            =   1290
         TabIndex        =   22
         Top             =   2220
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin VB.Shape shpCuadro 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00C00000&
         FillColor       =   &H00400000&
         Height          =   3195
         Index           =   0
         Left            =   105
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   6060
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
         Index           =   3
         Left            =   1650
         TabIndex        =   11
         Top             =   2265
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Devolución Quinta Categoria :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1980
         Width           =   2500
      End
      Begin VB.Label lblDato 
         Caption         =   "Remuneración  Anterior Empresa :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2500
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
         Left            =   1650
         TabIndex        =   5
         Top             =   1125
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
         Left            =   1650
         TabIndex        =   14
         Top             =   2835
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Tabla de UIT :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   2550
         Width           =   2500
      End
      Begin VB.Label lblDato 
         Caption         =   "Renta de Quinta Categoria :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1410
         Width           =   2500
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
         Left            =   1650
         TabIndex        =   8
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
         Index           =   0
         Left            =   1650
         TabIndex        =   2
         Top             =   555
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Remuneración  Ganada :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   270
         Width           =   2500
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
         Left            =   5790
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
         Picture         =   "prmcerti5ta.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   5400
         TabIndex        =   26
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmcerti5ta.frx":0044
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
         Left            =   285
         TabIndex        =   16
         Top             =   120
         Width           =   4800
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   4350
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   24
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
Attribute VB_Name = "fPrmCertifik5ta"
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
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Sub ShowScreen()
    
  ' Información de configuración
  s_Sql = "SELECT remganada, remanterior, codcpc5ta, codcpc5ta_ing, codtbluit "
  s_Sql = s_Sql & "FROM plcfgempresa "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    Me.Tag = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtRemunera(0), gdl_Funcion.aTexto(porstRecordset!remganada), Me.Tag, False, porstRecordset!remganada.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(1), gdl_Funcion.aTexto(porstRecordset!remanterior), Me.Tag, False, porstRecordset!remanterior.DefinedSize
    gdl_Procedure.EditText "AT", txtQuinta(0), gdl_Funcion.aTexto(porstRecordset!codcpc5ta), Me.Tag, False, porstRecordset!codcpc5ta.DefinedSize
    gdl_Procedure.EditText "AT", txtQuinta(1), gdl_Funcion.aTexto(porstRecordset!codcpc5ta_ing), Me.Tag, False, porstRecordset!codcpc5ta_ing.DefinedSize
    gdl_Procedure.EditText "AT", txtUit, gdl_Funcion.aTexto(porstRecordset!codtbluit), Me.Tag, False, porstRecordset!codtbluit.DefinedSize
  Else
    Me.Tag = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtRemunera(0), "", Me.Tag, False, porstRecordset!remganada.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(1), "", Me.Tag, False, porstRecordset!remanterior.DefinedSize
    gdl_Procedure.EditText "AT", txtQuinta(0), "", Me.Tag, False, porstRecordset!codcpc5ta.DefinedSize
    gdl_Procedure.EditText "AT", txtQuinta(1), "", Me.Tag, False, porstRecordset!codcpc5ta_ing.DefinedSize
    gdl_Procedure.EditText "AT", txtUit, "", Me.Tag, False, porstRecordset!codtbluit.DefinedSize
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(0), "CP")
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(1), "CP")
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtQuinta(0), "CP")
  lblHelp(3) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtQuinta(1), "CP")
  lblHelp(4) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, ps_Anyo & txtUit, "TB")

End Sub
']
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0, 1, 3   ' Conceptos de planilla remuneracion
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Ingresos"
    s_Registro = ps_ClsPlanilla & "F" & "0"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
   Case 2    ' Conceptos de planilla quinta
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Descuentos"
    s_Registro = ps_ClsPlanilla & "F" & "1"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
   Case 4     ' Tabla de uit
    tdbHelp.Columns(0).DataField = "codtbl": tdbHelp.Columns(1).DataField = "destbl"
    tdbHelp.Caption = "Tablas Generales"
    s_Sql = gdl_Funcion.HelpTablas("tbl", "codtbl", ps_Anyo & ps_ClsPlanilla, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + (Choose(Index + 1, cmdHelp(Index).Top, cmdHelp(Index).Top, 1400, 1500, 1600) + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  
  ' Realizo las validaciones de los campos a actualizar
  If txtRemunera(0) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Ganada", vbExclamation: txtRemunera(0).SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Concepto de Remuneración Ganada no es valido; Verificar", vbExclamation: txtRemunera(0).SetFocus: Exit Sub
  If txtRemunera(1) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Ganada Otra Empresa", vbExclamation: txtRemunera(1).SetFocus: Exit Sub
  If lblHelp(1) = "???" Then Beep: MsgBox "Concepto de Remuneración Ganada Otra Empresa no es valido; Verificar", vbExclamation: txtRemunera(1).SetFocus: Exit Sub
  If txtQuinta(0) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Renta Quinta Categoria", vbExclamation: txtQuinta(0).SetFocus: Exit Sub
  If lblHelp(2) = "???" Then Beep: MsgBox "Concepto de Renta Quinta Categoria no es valido; Verificar", vbExclamation: txtQuinta(0).SetFocus: Exit Sub
  If (txtQuinta(1) <> "" And lblHelp(4) = "???") Then Beep: MsgBox "Concepto de Devolución Renta Quinta Categoria no es valido; Verificar", vbExclamation: txtQuinta(1).SetFocus: Exit Sub
  If txtUit = "" Then Beep: MsgBox "Debe Ingresar el Codigo de Tabla de UIT", vbExclamation: txtUit.SetFocus: Exit Sub
  If lblHelp(4) = "???" Then Beep: MsgBox "Tabla de UIT no es valido; Verificar", vbExclamation: txtUit.SetFocus: Exit Sub
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("pdoano", "remganada", "remanterior", "codcpc5ta", "codcpc5ta_ing", "codtbluit", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, txtRemunera(0), txtRemunera(1), txtQuinta(0), txtQuinta(1), txtUit, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plcfgempresa", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plcfgempresa", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Ubico el registro ingresado o actualizado
  ShowScreen
  txtRemunera(0).SetFocus
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 5340: Me.Width = 6510
  Me.Left = 3580: Me.Top = 2500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Parametros de Certificado"
  lblTitle = "Parametros"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "edit": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "aceptar"
  aElemento(0, 2) = "Actualizar Información de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdUpdate, aElemento
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
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
  gdl_Procedure.InicializaGrilla tdbHelp, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Cálculo", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0, 1      ' Concepto de remuneración ganada
    txtRemunera(n_IndexHelp) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtRemunera(n_IndexHelp).SetFocus
   Case 2, 3      ' Concepto de quinta categoria
    txtQuinta(n_IndexHelp - 2) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtQuinta(n_IndexHelp - 2).SetFocus
   Case 3       ' Entidad bancaria
    txtUit = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtUit.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0, 1, 3   ' Conceptos de remuneración ganada
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 2     ' Concepto de quinta categoria
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Act
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 3     ' Tabla de UIT
    s_Sql = gdl_Funcion.HelpTablas("tbl", tdbHelp.Columns(ColIndex).DataField, ps_Anyo & ps_ClsPlanilla, "")
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
Private Sub txtQuinta_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtQuinta(Index)
End Sub
Private Sub txtQuinta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 2
End Sub
Private Sub txtQuinta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtQuinta_LostFocus(Index As Integer)
  lblHelp(Index + 2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtQuinta(Index), "CP")
End Sub
Private Sub txtRemunera_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtRemunera(Index)
End Sub
Private Sub txtRemunera_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtRemunera_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_LostFocus(Index As Integer)
  lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(Index), "CP")
End Sub
Private Sub txtUit_GotFocus()
  gdl_Procedure.MarcaGet txtUit
End Sub
Private Sub txtUit_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 4
End Sub
Private Sub txtUit_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtUit_LostFocus()
  lblHelp(4) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, ps_Anyo & txtUit, "TB")
End Sub
