VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fContabilizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   Icon            =   "contabiliza.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   9105
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   25
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
      Begin VB.ComboBox cboParametro 
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
         ItemData        =   "contabiliza.frx":000C
         Left            =   1335
         List            =   "contabiliza.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   105
         Width           =   3165
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Formato :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   510
         TabIndex        =   26
         Top             =   150
         Width           =   720
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   1
         Left            =   330
         Shape           =   4  'Rounded Rectangle
         Top             =   45
         Width           =   4425
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   3255
      Index           =   1
      Left            =   15
      TabIndex        =   18
      Top             =   3270
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   5741
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
         Height          =   3100
         Left            =   45
         TabIndex        =   19
         Top             =   120
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   5477
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
         BorderStyle     =   0
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
         CollapseColor   =   12632064
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&"
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
   Begin Threed.SSFrame frmCuadro 
      Height          =   2700
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   540
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   4762
      _StockProps     =   14
      Caption         =   " Parametro de Contabilización "
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
      Begin VB.ComboBox cboSeccion 
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
         ItemData        =   "contabiliza.frx":0010
         Left            =   1455
         List            =   "contabiliza.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1755
         Width           =   4110
      End
      Begin VB.ComboBox cboPeriodo 
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
         ItemData        =   "contabiliza.frx":0014
         Left            =   1455
         List            =   "contabiliza.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1395
         Width           =   4110
      End
      Begin VB.ComboBox cmbProceso 
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
         ItemData        =   "contabiliza.frx":0018
         Left            =   1455
         List            =   "contabiliza.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2115
         Width           =   4110
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
         ItemData        =   "contabiliza.frx":001C
         Left            =   1455
         List            =   "contabiliza.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1035
         Width           =   2655
      End
      Begin VB.TextBox txtComprobante 
         ForeColor       =   &H00C00000&
         Height          =   280
         Left            =   6795
         TabIndex        =   4
         Text            =   "999999"
         Top             =   375
         Width           =   675
      End
      Begin VB.TextBox txtDiario 
         ForeColor       =   &H00C00000&
         Height          =   280
         Left            =   1455
         TabIndex        =   2
         Text            =   "9999"
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox txtGlosa 
         ForeColor       =   &H00C00000&
         Height          =   280
         Left            =   1455
         TabIndex        =   6
         Top             =   705
         Width           =   6015
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   6180
         TabIndex        =   10
         Top             =   1035
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   140836865
         CurrentDate     =   37515
      End
      Begin Threed.SSCheck chkProceso 
         Height          =   285
         Left            =   5715
         TabIndex        =   17
         Top             =   2130
         Width           =   960
         _Version        =   65536
         _ExtentX        =   1693
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "General"
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
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proceso :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   405
         TabIndex        =   15
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   11
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sección :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   405
         TabIndex        =   13
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mes :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   7
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   5130
         TabIndex        =   9
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comprobante :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   5625
         TabIndex        =   3
         Top             =   375
         Width           =   1050
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diario :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   1
         Top             =   375
         Width           =   930
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Glosa :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   405
         TabIndex        =   5
         Top             =   735
         Width           =   930
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2280
         Index           =   0
         Left            =   300
         Shape           =   4  'Rounded Rectangle
         Top             =   270
         Width           =   7635
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   20
      Top             =   6600
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
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   5850
         TabIndex        =   23
         Top             =   105
         Width           =   1140
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   6975
         TabIndex        =   24
         Top             =   105
         Width           =   1140
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   3585
         TabIndex        =   21
         Top             =   105
         Width           =   1140
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   4710
         TabIndex        =   22
         Top             =   105
         Width           =   1140
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   5940
      Index           =   0
      Left            =   8280
      TabIndex        =   28
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   10477
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
         TabIndex        =   31
         Tag             =   "0"
         Top             =   1020
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
         Picture         =   "contabiliza.frx":0020
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   29
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
         TabIndex        =   30
         Tag             =   "0"
         Top             =   585
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
         Picture         =   "contabiliza.frx":003C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   33
         Tag             =   "0"
         Top             =   2190
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
         Picture         =   "contabiliza.frx":0058
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   34
         Tag             =   "0"
         Top             =   2625
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
         Picture         =   "contabiliza.frx":0074
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   32
         Tag             =   "0"
         Top             =   1455
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
         Picture         =   "contabiliza.frx":0090
      End
   End
End
Attribute VB_Name = "fContabilizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private n_Index As Integer                              ' Indice de la barra de herramientas, indice para bucle
Private s_OptRegistro As String                         ' Instancia del formulario activo
Private s_ConexiConta As String                         ' Cadena conexion contabilidad
Private s_StatusValid_DatosConcar As String             'Indicador ok/NO_OK las validaciones

Public Sub GenArchivoPLCentroExcel_V1(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim poApplExcel As Object, poLibroExcel As Object
  Dim sHojaExcel As String, sMoneda As String
  Dim sExpresion As String, s_OldMessage As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nSecuencia As Long
  Dim dFechaTCambio As Date
  
  Dim NAcum_Debe As Double
  Dim NAcum_Haber As Double
 
  ' Inicializando variable status validaciones
  s_StatusValid_DatosConcar = "OK"
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo
  
  nTipoCambio = 1
  dFechaTCambio = Format(dtpFecha, s_FmtFechMysql_0)
  ' Obtengo el tipo de cambio
  s_Sql = "SELECT codpdo, tipocambio,fechaproceso "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
    dFechaTCambio = porstRecordset!fechaproceso
  End If
  porstRecordset.Close
  
  ' Recupero la información para exportar
  s_Sql = "SELECT tmpx.codcta, tmpx.codpsn, "
  s_Sql = s_Sql & "CASE WHEN debe_mn>0 THEN 'DEBE' ELSE 'HABER' END AS tipoDH, "
  s_Sql = s_Sql & "tmpx.debe_mn,tmpx.haber_mn, 'P0' AS IND, "
  s_Sql = s_Sql & "CASE WHEN (tmpx.codcco='RA101' OR tmpx.codcco='RA105' OR tmpx.codcco='RA109' OR tmpx.codcco='RR100' OR tmpx.codcco='RC101') THEN 'PE00' ELSE RIGHT(tmpx.codcco,4) END CentroCosto, "
  s_Sql = s_Sql & "Case WHEN "
  s_Sql = s_Sql & "CASE WHEN (tmpx.codcco='RA101' OR tmpx.codcco='RA105' OR tmpx.codcco='RA109' OR tmpx.codcco='RR100' OR tmpx.codcco='RC101') THEN 'PE00' ELSE RIGHT(tmpx.codcco,4) END"
       s_Sql = s_Sql & "= 'PE00' THEN tmpx.codcco "
       s_Sql = s_Sql & "Else "
       s_Sql = s_Sql & " @SUBSIDIARIA:=CASE RIGHT(tmpx.codcco,4) "
       s_Sql = s_Sql & "WHEN  'PE01' THEN 'RPER101101' "
       s_Sql = s_Sql & "WHEN  'PE03' THEN 'RPER101301' "
       s_Sql = s_Sql & "WHEN  'PE06' THEN 'RPER101601' "
       s_Sql = s_Sql & "WHEN  'PE08' THEN 'RPER101801' "
       s_Sql = s_Sql & "WHEN  'PE10' THEN 'RPER102001' "
       s_Sql = s_Sql & "WHEN  'PE11' THEN 'RPER102101' "
       s_Sql = s_Sql & "WHEN  'PE16' THEN 'RPER102601' ELSE '-' End "
       s_Sql = s_Sql & "END subsidiaria, "
       s_Sql = s_Sql & "CONCAT(LEFT(tmpx.detalle,21),'.') AS Glosa, CONCAT(LEFT(tmpx.detalle,21),'.') AS Texto,tmpx.codmon,"
       s_Sql = s_Sql & "CASE WHEN debe_mn>0 THEN '1' ELSE '2' END AS lista "
       s_Sql = s_Sql & "FROM " & s_Archivo & " tmpx "
       s_Sql = s_Sql & "ORDER BY LISTA"
  
  
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
'      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
'      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
       a_Campos = Array("codcta", "codpsn", "tipoDH", "debe_mn", "haber_mn", "ind", "CentroCosto", "subsidiaria", "Glosa", "texto", "codmon", "lista")
       a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    Else
      ' Creo objeto de archivo
      Set poApplExcel = CreateObject("Excel.Application")
      poApplExcel.Visible = False
      sExpresion = Trim(cmbPeriodo.Text)
      Set poLibroExcel = poApplExcel.Workbooks.Add
      sHojaExcel = Left(sExpresion, 20)
      poLibroExcel.Sheets("Hoja1").Name = sHojaExcel
      
      'nSecuencia = 1
      nSecuencia = 10
      ' Titulos de registro
      sExpresion = "CUENTA"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
      sExpresion = " "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
      
      sExpresion = "TIPO_CUENTA"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
      
      sExpresion = "DEBE"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
      
      sExpresion = "HABER"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
      sExpresion = "IND."
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
      sExpresion = "C.COSTO "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
      
      sExpresion = " "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
      
      sExpresion = "GLOSA"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
      sExpresion = "TEXTO "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
      
    End If
    'nSecuencia = 2
    nSecuencia = 11
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        'a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        '"codcta", "codpsn", "tipoDH", "debe_mn", "haber_mn", "ind", "CentroCosto", "subsidiaria", "Glosa", "texto", "codmon", "lista")
        a_Valores = Array(gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("tipoDH")), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), gdl_Funcion.aTexto(porstRecordset("ind")), gdl_Funcion.aTexto(porstRecordset("CentroCosto")), gdl_Funcion.aTexto(porstRecordset("Subsidiaria")), gdl_Funcion.aTexto(porstRecordset("Glosa")), gdl_Funcion.aTexto(porstRecordset("texto")), gdl_Funcion.aTexto(porstRecordset("codmon")), gdl_Funcion.aTexto(porstRecordset("lista")))
        
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        ' detalle por moneda
        sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
        If porstRecordset!codmon = sMoneda Then
        
          'CUENTA
          sExpresion = porstRecordset!codcta
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
          ' CODPSN
          sExpresion = "  "
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
          ' TIPO CUENTA
          sExpresion = porstRecordset!tipoDH
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
          
           ' DEBE MN
          sExpresion = porstRecordset!debe_mn
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
           ' HABER MN
          sExpresion = porstRecordset!haber_mn
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
          
           ' IND
          sExpresion = porstRecordset!ind
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
          
           ' CENTRO COSTO
          sExpresion = IIf(IsNull(porstRecordset!CentroCosto) = True, " ", porstRecordset!CentroCosto)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
          
           ' SUBSIDIARIA
          sExpresion = porstRecordset!subsidiaria
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
          
           ' GLOSA
          sExpresion = porstRecordset!glosa
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
          
           ' TEXTO
          sExpresion = porstRecordset!texto
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
                    
        End If
      End If
      
      NAcum_Debe = NAcum_Debe + porstRecordset!debe_mn
      NAcum_Haber = NAcum_Haber + porstRecordset!haber_mn
      
      
      ' Incremento el porcentaje
      nSecuencia = nSecuencia + 1
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    'Mostrando Sumatoria Total D/H
    poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "#,##0.00"
    poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = NAcum_Debe
    
     poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "#,##0.00"
    poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = NAcum_Haber
    
    If s_Accion = "G" Then
      ' Cierro y grabo documento excel
      sExpresion = Strings.Right(s_File, 4)
      If sExpresion = ".xls" Then
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlExcel8
      Else
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlOpenXMLWorkbook
         
      End If
      poLibroExcel.Close SaveChanges:=False
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Saco de memoria objeto
  Set poLibroExcel = Nothing
  Set poApplExcel = Nothing
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  

End Sub


Public Sub GenArchivoPLCentroExcel(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim poApplExcel As Object, poLibroExcel As Object
  Dim sHojaExcel As String, sMoneda As String
  Dim sExpresion As String, s_OldMessage As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nSecuencia As Long
  Dim dFechaTCambio As Date
  
  Dim NAcum_Debe As Double
  Dim NAcum_Haber As Double
 
  ' Inicializando variable status validaciones
  s_StatusValid_DatosConcar = "OK"
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo
  
  nTipoCambio = 1
  dFechaTCambio = Format(dtpFecha, s_FmtFechMysql_0)
  ' Obtengo el tipo de cambio
  s_Sql = "SELECT codpdo, tipocambio,fechaproceso "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
    dFechaTCambio = porstRecordset!fechaproceso
  End If
  porstRecordset.Close
  
  ' Recupero la información para exportar
  s_Sql = "SELECT tmpx.codcta, tmpx.codpsn, "
  s_Sql = s_Sql & "CASE WHEN debe_mn>0 THEN 'DEBE' ELSE 'HABER' END AS tipoDH, "
  s_Sql = s_Sql & "tmpx.debe_mn,tmpx.haber_mn, 'P0' AS IND, "
  s_Sql = s_Sql & "CASE WHEN (tmpx.codcco='RA101' OR tmpx.codcco='RA105' OR tmpx.codcco='RA109' OR tmpx.codcco='RR100' OR tmpx.codcco='RC101') THEN 'PE00' ELSE RIGHT(tmpx.codcco,4) END CentroCosto, "
  s_Sql = s_Sql & "Case WHEN "
  s_Sql = s_Sql & "CASE WHEN (tmpx.codcco='RA101' OR tmpx.codcco='RA105' OR tmpx.codcco='RA109' OR tmpx.codcco='RR100' OR tmpx.codcco='RC101') THEN 'PE00' ELSE RIGHT(tmpx.codcco,4) END"
       s_Sql = s_Sql & "= 'PE00' THEN tmpx.codcco "
       s_Sql = s_Sql & "Else "
       s_Sql = s_Sql & " @SUBSIDIARIA:=CASE RIGHT(tmpx.codcco,4) "
       s_Sql = s_Sql & "WHEN  'PE01' THEN 'RPER101101' "
       s_Sql = s_Sql & "WHEN  'PE03' THEN 'RPER101301' "
       s_Sql = s_Sql & "WHEN  'PE06' THEN 'RPER101601' "
       s_Sql = s_Sql & "WHEN  'PE08' THEN 'RPER101801' "
       s_Sql = s_Sql & "WHEN  'PE10' THEN 'RPER102001' "
       s_Sql = s_Sql & "WHEN  'PE11' THEN 'RPER102101' "
       s_Sql = s_Sql & "WHEN  'PE16' THEN 'RPER102601' ELSE '-' End "
       s_Sql = s_Sql & "END subsidiaria, "
       s_Sql = s_Sql & "CONCAT(LEFT(tmpx.detalle,21),'.') AS Glosa, CONCAT(LEFT(tmpx.detalle,21),'.') AS Texto,tmpx.codmon,"
       s_Sql = s_Sql & "CASE WHEN debe_mn>0 THEN '1' ELSE '2' END AS lista "
       s_Sql = s_Sql & "FROM " & s_Archivo & " tmpx "
       s_Sql = s_Sql & "ORDER BY LISTA"
  
  
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
'      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
'      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
       a_Campos = Array("codcta", "codpsn", "tipoDH", "debe_mn", "haber_mn", "ind", "CentroCosto", "subsidiaria", "Glosa", "texto", "codmon", "lista")
       a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    Else
      ' Creo objeto de archivo
      Set poApplExcel = CreateObject("Excel.Application")
      poApplExcel.Visible = False
      sExpresion = Trim(cmbPeriodo.Text)
      Set poLibroExcel = poApplExcel.Workbooks.Add
      sHojaExcel = Left(sExpresion, 20)
      poLibroExcel.Sheets("Hoja1").Name = sHojaExcel
      
      
      
      
      'nSecuencia = 1
      nSecuencia = 10
      ' Titulos de registro
      sExpresion = "NRO DOC"
      poLibroExcel.Sheets(sHojaExcel).Cells(9, 1).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(9, 1).Value = sExpresion
      
      sExpresion = txtDiario.Text & txtComprobante.Text
      poLibroExcel.Sheets(sHojaExcel).Cells(9, 2).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(9, 2).Value = sExpresion
      
      sExpresion = "CUENTA"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
      sExpresion = " "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
      
      sExpresion = "TIPO_CUENTA"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
      
      sExpresion = "IMPORTE" 'debe
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
      
       sExpresion = " "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
      
      sExpresion = " " 'haber
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
      sExpresion = "ASIGNACION"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
      
      sExpresion = " "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
      sExpresion = "TEXTO"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
      
     
      
      sExpresion = "DIV."
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
      sExpresion = "C.COSTO"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
      
    End If
    'nSecuencia = 2
    nSecuencia = 11
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        'a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        '"codcta", "codpsn", "tipoDH", "debe_mn", "haber_mn", "ind", "CentroCosto", "subsidiaria", "Glosa", "texto", "codmon", "lista")
        a_Valores = Array(gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("tipoDH")), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), gdl_Funcion.aTexto(porstRecordset("ind")), gdl_Funcion.aTexto(porstRecordset("CentroCosto")), gdl_Funcion.aTexto(porstRecordset("Subsidiaria")), gdl_Funcion.aTexto(porstRecordset("Glosa")), gdl_Funcion.aTexto(porstRecordset("texto")), gdl_Funcion.aTexto(porstRecordset("codmon")), gdl_Funcion.aTexto(porstRecordset("lista")))
        
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        ' detalle por moneda
        sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
        If porstRecordset!codmon = sMoneda Then
        
          'CUENTA
          sExpresion = porstRecordset!codcta
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
          ' CODPSN
          sExpresion = "  "
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
          ' TIPO CUENTA
          sExpresion = porstRecordset!tipoDH
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
          
           ' IMPORTE "D"
          sExpresion = porstRecordset!debe_mn
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
           
           ' IND
          sExpresion = porstRecordset!ind
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
           
           ' HABER MN
          sExpresion = porstRecordset!haber_mn
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
          
           ' GLOSA
          sExpresion = porstRecordset!glosa
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
          
          sExpresion = " "
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
          
           ' TEXTO
          sExpresion = porstRecordset!texto & " " & Me.dtpFecha
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
          
           ' DIVISIONARIA
          sExpresion = porstRecordset!subsidiaria
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
          
           ' CENTRO COSTO
          sExpresion = IIf(IsNull(porstRecordset!CentroCosto) = True, " ", porstRecordset!CentroCosto)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
          
          
          
                    
        End If
      End If
      
      NAcum_Debe = NAcum_Debe + porstRecordset!debe_mn
      NAcum_Haber = NAcum_Haber + porstRecordset!haber_mn
      
      
      ' Incremento el porcentaje
      nSecuencia = nSecuencia + 1
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    'Mostrando Sumatoria Total D/H
    poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "#,###,##0.00"
    poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = NAcum_Debe
    
     poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "#,###,##0.00"
    poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = NAcum_Haber
    
    If s_Accion = "G" Then
      ' Cierro y grabo documento excel
      sExpresion = Strings.Right(s_File, 4)
      If sExpresion = ".xls" Then
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlExcel8
      Else
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlOpenXMLWorkbook
         
      End If
      poLibroExcel.Close SaveChanges:=False
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Saco de memoria objeto
  Set poLibroExcel = Nothing
  Set poApplExcel = Nothing
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  

End Sub


'[
Private Sub ContabilizaComprobante(ByVal s_Archivo As String)
  Dim n_Importe As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim sExpresion As String, s_OldMessage As String
  Dim sConexionStr As String
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo

  ' Obtengo el tipo de cambio
  nTipoCambio = 1
  s_Sql = "SELECT codpdo, tipocambio "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
  End If
  porstRecordset.Close

  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  '[ Inicio la conexión a la base de datos ]
  sConexionStr = OpenConnection(ps_ServidorCon, "sysmacon")
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Comprobante ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    
    ' Genero el registro de grabación
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción

    ' Cabecera de comprobante
    a_Campos = Array("codemp", "pdoano", "mespvs", "coddro", "nrocpb", "fehcpb", "glocpb", "tpognr", "indncu", "indanu", "usrcre", "fyhcre")
    a_Valores = Array(ps_EmpresaCon, ps_Anyo, Left(cmbPeriodo.Text, 2), Trim(txtDiario.Text), gdl_Funcion.PadL(txtComprobante.Text, 6, "0"), Format(dtpFecha.Value, s_FmtFechMysql_0), Trim(txtGlosa.Text), s_Estado_Ina, s_Estado_Ina, s_Estado_Ina, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
    If Not Records_Ins("cocpbcab", a_Campos, a_Valores, a_Tipos) Then GoTo Error
    While Not porstRecordset.EOF
      nRegistro = nRegistro + 1
      ' Realizo la actualización de los registros
      sExpresion = ps_Anyo & "-" & Left(Trim(cmbPeriodo.Text), 2)
      sExpresion = gdl_Funcion.PadL(IIf(cboPeriodo.ListIndex <> 0, Trim(Left(cboPeriodo.Text, 8)), sExpresion), 10, "0")
      sExpresion = IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, sExpresion, "")
      a_Campos = Array("codemp", "pdoano", "mespvs", "coddro", "nrocpb", "nroite", "blqite", "codtdc", "fehope", "codcta", "codcco", "codaux", "serdoc", "nrodoc", "feedoc", "fevdoc", "ferdoc", "refdoc", "pdocpr", "gloite", "gloitex", "tpoctb", "tpopvs", "tpomon", "tpotcb", "imptcb", "impmn", "impme", "tpognr", "indfjo_det", "indgnr_rp", "tpodoc", "codcon", "usrcre", "fyhcre")
      a_Valores = Array(ps_EmpresaCon, ps_Anyo, Left(cmbPeriodo.Text, 2), Trim(txtDiario.Text), gdl_Funcion.PadL(txtComprobante.Text, 6, "0"), nRegistro, nRegistro, IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, "00", ""), Format(dtpFecha.Value, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!codcta), gdl_Funcion.aTexto(porstRecordset!codcco), gdl_Funcion.aTexto(porstRecordset!codpsn), _
                        IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, "PLLA", ""), sExpresion, Format(dtpFecha.Value, s_FmtFechMysql_0), Format(dtpFecha.Value, s_FmtFechMysql_0), Format(dtpFecha.Value, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!codref), "", gdl_Funcion.aTexto(porstRecordset!detalle), "", IIf(porstRecordset!debe_mn > 0, "D", "H"), IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, "P", "O"), gdl_Funcion.aTexto(porstRecordset!codmon), _
                        gdl_Funcion.aTexto(porstRecordset!tpotcb), Format(nTipoCambio, s_FormatoNum_1), Round(CDec(porstRecordset!debe_mn) + CDec(porstRecordset!haber_mn), 2), Round(CDec(porstRecordset!debe_me) + CDec(porstRecordset!haber_me), 2), s_Estado_Ina, s_Estado_Ina, s_Estado_Ina, "", "", ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      If Not Records_Ins("cocpbdet", a_Campos, a_Valores, a_Tipos) Then GoTo Error
      ' Incremento el porcentaje
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
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
Private Sub ExportaContabilidad(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String
  Dim n_Importe As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim sExpresion As String, s_OldMessage As String
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo

  nTipoCambio = 1
  ' Obtengo el tipo de cambio
  s_Sql = "SELECT codpdo, tipocambio "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
  End If
  porstRecordset.Close

  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_File, True)
      s_Caracter = "|"
    End If
    While Not porstRecordset.EOF
      nRegistro = nRegistro + 1
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = ""
        psRegistro = psRegistro & ps_Anyo & s_Caracter
        psRegistro = psRegistro & Trim(txtDiario.Text) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.PadL(txtComprobante.Text, 6, "0") & s_Caracter
        psRegistro = psRegistro & Left(Trim(cmbPeriodo.Text), 2) & s_Caracter
        psRegistro = psRegistro & Format(dtpFecha, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & Trim(txtGlosa.Text) & s_Caracter
        psRegistro = psRegistro & Space(1) & s_Caracter
        psRegistro = psRegistro & "D" & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.PadL(nRegistro, 4, "0") & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.PadL(nRegistro, 4, "0") & s_Caracter
        psRegistro = psRegistro & IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, "00", "") & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codcta) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codcco) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codpsn) & s_Caracter
        psRegistro = psRegistro & IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, "PLLA", "") & s_Caracter
        sExpresion = ps_Anyo & "-" & Left(Trim(cmbPeriodo.Text), 2)
        sExpresion = gdl_Funcion.PadL(IIf(cboPeriodo.ListIndex <> 0, Trim(Left(cboPeriodo.Text, 8)), sExpresion), 10, "0")
        psRegistro = psRegistro & IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, sExpresion, "") & s_Caracter
        psRegistro = psRegistro & Format(dtpFecha, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & Format(dtpFecha, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & Format(dtpFecha, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codref) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!detalle) & s_Caracter
        psRegistro = psRegistro & Space(1) & s_Caracter
        psRegistro = psRegistro & IIf(porstRecordset!debe_mn > 0, "D", "H") & s_Caracter
        psRegistro = psRegistro & IIf(gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act, "P", "O") & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codmon) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!tpotcb) & s_Caracter
        psRegistro = psRegistro & Format(nTipoCambio, s_FormatoNum_1) & s_Caracter
        n_Importe = CDec(porstRecordset!debe_mn) + CDec(porstRecordset!haber_mn)
        psRegistro = psRegistro & Format(n_Importe, "###########0.00") & s_Caracter
        n_Importe = CDec(porstRecordset!debe_me) + CDec(porstRecordset!haber_me)
        psRegistro = psRegistro & Format(n_Importe, "###########0.00") & s_Caracter
        psRegistro = psRegistro & "" & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.PadL(nRegistro, 4, "0") & s_Caracter
        potxtFileExp.WriteLine psRegistro
      End If
      ' Incremento el porcentaje
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    If s_Accion = "G" Then
      ' Cierro objeto y saco de memoria
      potxtFileExp.Close
    End If
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
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
Public Sub GenArchivoConcarExcel(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim poApplExcel As Object, poLibroExcel As Object
  Dim sHojaExcel As String, sMoneda As String
  Dim sExpresion As String, s_OldMessage As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nSecuencia As Long
  Dim dFechaTCambio As Date
  
  'Aplicando las validaciones de ingreso de Datos
  
  If Len(txtDiario.Text) <> 2 Then MsgBox "El sub diario debe ser de 2 caracteres de longitud", vbCritical: s_StatusValid_DatosConcar = "NO_OK": Exit Sub
  If Len(txtComprobante.Text) <> 6 Then MsgBox "El Numero de comprobante debe ser de 6 caracteres de longitud", vbCritical: s_StatusValid_DatosConcar = "NO_OK": Exit Sub
  
  If IsNumeric(Left(txtComprobante.Text, 2)) = True Then
      If Val(Left(txtComprobante.Text, 2)) < 1 Or Val(Left(txtComprobante.Text, 2)) > 12 Then
         MsgBox "Los dos primeros Digitos del comprobantes deben tener un valor de mes valido", vbCritical
         s_StatusValid_DatosConcar = "NO_OK"
         Exit Sub
      End If
  Else
    MsgBox "Los dos primeros Digitos del comprobantes deben tener un valor numérico", vbInformation
    s_StatusValid_DatosConcar = "NO_OK"
    Exit Sub
  End If
  
  If Len(txtGlosa.Text) > 40 Then MsgBox "El numero maximo de caracteres de la glosa no debe ser mayor a 40 caracteres", vbCritical: s_StatusValid_DatosConcar = "NO_OK": Exit Sub

  ' Inicializando variable status validaciones
  s_StatusValid_DatosConcar = "OK"
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo
  
  nTipoCambio = 1
  dFechaTCambio = Format(dtpFecha, s_FmtFechMysql_0)
  ' Obtengo el tipo de cambio
  s_Sql = "SELECT codpdo, tipocambio, fechaproceso "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
    dFechaTCambio = porstRecordset!fechaproceso
  End If
  porstRecordset.Close
  
  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me, "
  s_Sql = s_Sql & "IFNULL(cco.detcco,cta.detcta) AS detalleitem "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco AND cco.estcco='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set poApplExcel = CreateObject("Excel.Application")
      poApplExcel.Visible = False
      sExpresion = Trim(cmbPeriodo.Text)
      Set poLibroExcel = poApplExcel.Workbooks.Add
      sHojaExcel = Left(sExpresion, 20)
      poLibroExcel.Sheets("Hoja1").Name = sHojaExcel
    
      nSecuencia = 1
      ' Titulos de registro
      sExpresion = "Sub Diario"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
      
      sExpresion = "Numero de Comprobante"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
      
      sExpresion = "Fecha de Comprobante"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
      
      sExpresion = "Codigo de Moneda"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
      
      sExpresion = "Glosa Principal"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
      
      sExpresion = "Tipo de Cambio"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
      
      sExpresion = "Tipo de Conversion"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
      
      sExpresion = "Flag de Conversion de Moneda"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
      
      sExpresion = "Fecha Tipo de Cambio"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
      
      sExpresion = "Cuenta Contable"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
      
      sExpresion = "Codigo de Anexo"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
      
      sExpresion = "Codigo Centro de Costo"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).Value = sExpresion
      
      sExpresion = "Debe/Haber"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).Value = sExpresion
      
      sExpresion = "Importe Original"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).Value = sExpresion
      
      sExpresion = "Importe en " & s_Codmon_me_Nom
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).Value = sExpresion
      
      sExpresion = "Importe en " & s_Codmon_mn_Nom
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).Value = sExpresion
      
      
      sExpresion = "Tipo de Documento"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).Value = sExpresion
      
      sExpresion = "Numero de Documento"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).Value = sExpresion
      
      sExpresion = "Fecha de Documento"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).Value = sExpresion
      
      sExpresion = "Fecha de Vencimiento"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).Value = sExpresion
      
      sExpresion = "Codigo de Area"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).Value = sExpresion
      
      sExpresion = "Glosa Detalle"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).Value = sExpresion
      
      sExpresion = "Codigo de Anexo Auxiliar"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).Value = sExpresion
      
      sExpresion = "Medio de Pago"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).Value = sExpresion
      
      sExpresion = "Tipo de Documento Refeferencia"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).Value = sExpresion
      
      sExpresion = "Numero de Documento de Referencia"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).Value = sExpresion
      
      sExpresion = "Fecha de Documento de Referencia"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).Value = sExpresion
      
      sExpresion = "Base Imponible de Documento de Referencia"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).Value = sExpresion
      
      sExpresion = "IGV Documento Provision"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).Value = sExpresion
      
      sExpresion = "Tipo Referencia en estado MQ"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).Value = sExpresion
      
      sExpresion = "Numero Serie Caja Registradora"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).Value = sExpresion
      
      sExpresion = "Fecha de Operacion"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).Value = sExpresion
      
      sExpresion = "Tipo de Tasa"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).Value = sExpresion
      
      sExpresion = "Tasa Detraccion/Percepción"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).Value = sExpresion
      
      sExpresion = "Importe Base Detraccion/ Percepción " & s_Codmon_me_Nom
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).Value = sExpresion
      
      sExpresion = "Importe Base Detraccion/ Percepción " & s_Codmon_mn_Nom
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).Value = sExpresion
    End If
    
    nSecuencia = 2
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        ' detalle por moneda
        sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
        If porstRecordset!codmon = sMoneda Then
         ' Sub Diario
          sExpresion = Trim(txtDiario.Text)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
          ' Numero de Comprobante
          sExpresion = Trim(txtComprobante.Text)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
          ' Fecha de comprobante
          sExpresion = Format(dtpFecha.Value, "dd/mm/yyyy")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
          ' Codigo de Moneda
          sMoneda = IIf(porstRecordset!codmon = s_Codmon_mn, "MN", "US")
          sExpresion = sMoneda
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
          ' Glosa Principal
          sExpresion = Trim(txtGlosa.Text)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
          ' Tipo de Cambio
          nImporte = Format(nTipoCambio, "####0.000000")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "####0.000000"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = nImporte
          ' Tipo de Conversion
          sExpresion = "F"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
          ' Flag de Conversion de Moneda
          sExpresion = "S"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
           'Fecha de tipo de Cambio
           sExpresion = Format(dFechaTCambio, "dd/mm/yyyy")
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
          
          'Cuenta Contable
          sExpresion = porstRecordset!codcta
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
          'Codigo de Anexo
          sExpresion = gdl_Funcion.aTexto(porstRecordset!codpsn)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
          'Codigo del Centro de Costo
          sExpresion = IIf(IsNull(porstRecordset!codcco) = True, "", porstRecordset!codcco)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).Value = sExpresion
          'Debe/Haber
          sExpresion = IIf(porstRecordset!debe_mn > 0, "D", "H")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).Value = sExpresion
          'Importe Original
          nImporte = CDec(porstRecordset("debe_m" & porstRecordset!codmon)) + CDec(porstRecordset("haber_m" & porstRecordset!codmon))
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).Value = nImporte
          'Importe Dolares
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).Value = sExpresion
          'Importe en Soles
          nImporte = CDec(porstRecordset("debe_m" & porstRecordset!codmon)) + CDec(porstRecordset("haber_m" & porstRecordset!codmon))
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).Value = nImporte
          'Tipo de Documento
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).Value = sExpresion
          'Numero de Documento
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).Value = sExpresion
          'Fecha de Documento
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).Value = sExpresion
          'Fecha de Vencimiento
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).Value = sExpresion
          'Codigo del Area
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).Value = sExpresion
          'Glosa Detalle
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).Value = sExpresion
          'Codigo del Anexo Auxiliar
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).Value = sExpresion
          'Medio de Pago
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).Value = sExpresion
          'Tipo de Documento de Referencia
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).Value = sExpresion
          'Numero de Documento de Referencia
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).Value = sExpresion
          'Fecha de Documento de Referencia
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).Value = sExpresion
          'Base imponible de Documento de Referencia
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).Value = sExpresion
          'IGV Documento de Provisión
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).Value = sExpresion
          'Tipo de Referencia en estado MQ
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).Value = sExpresion
         'Numero de seruie de Caja Registradora
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).Value = sExpresion
          'Fecha de Operación
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).Value = sExpresion
          'Tipo de Tasa
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).Value = sExpresion
          'Tasa Detraccion/Percepcion
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).Value = sExpresion
          'Importe Base Detracción/Percepción Dolares
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).Value = sExpresion
          'Importe Base Detracción/Percepción Soles
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).Value = sExpresion
        End If
      End If
      ' Incremento el porcentaje
      nSecuencia = nSecuencia + 1
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend

    If s_Accion = "G" Then
      ' Cierro y grabo documento excel
      sExpresion = Strings.Right(s_File, 4)
      If sExpresion = ".xls" Then
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlExcel8
      Else
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlOpenXMLWorkbook
      End If
      poLibroExcel.Close SaveChanges:=False
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Saco de memoria objeto
  Set poLibroExcel = Nothing
  Set poApplExcel = Nothing
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  

End Sub


Private Sub GenArchivoOracle(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String, sHojaInterface As String
  Dim n_Importe As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long, nSecuencia As Long
  Dim sExpresion As String, s_OldMessage As String
  Dim poAplicacionExcel As Object, poLibroExcelSave As Object
  
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo

  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.codsec, tmp.detalle, tmp.codmon, cfg.lineanegocio, cfg.segmentonego, "
  s_Sql = s_Sql & "cfg.clientenego, cta.tposdo, cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "LEFT JOIN plcfgcencosto cfg ON cfg.codcco=tmp.codcco "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set poAplicacionExcel = CreateObject("Excel.Application")
      poAplicacionExcel.Visible = False
      sHojaInterface = "Interfase Archivo"
      Set poLibroExcelSave = poAplicacionExcel.Workbooks.Add
      poLibroExcelSave.Sheets("Hoja1").Name = sHojaInterface
      nSecuencia = 1
      ' Titulos
      poLibroExcelSave.Sheets(sHojaInterface).Range(poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1), poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33)).Font.Name = "Arial"
      poLibroExcelSave.Sheets(sHojaInterface).Range(poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1), poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33)).Font.Size = 8
      poLibroExcelSave.Sheets(sHojaInterface).Range(poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1), poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33)).Font.Bold = True
      poLibroExcelSave.Sheets(sHojaInterface).Range(poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1), poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33)).Interior.TintAndShade = 0.599993896298105
      For n_Index = 1 To 33
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, n_Index).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, n_Index).Value = Choose(n_Index, "Consecutivo", "Estado", "Categoría", "Origen", "Nombre Lote", "Nombre Asiento", "Descripción Asiento", "Cia", "Cuenta", "Libre1", "Area", "Concepto", "Emp Grupo", "Cliente", "Servicio Proyectado", "Empleado", "Localización", "Libre2", "Descripción Línea", "Debito", "Credito", "Moneda", "Fecha Cambio Moneda", "Tipo", "Periodo_Liquidado", "Usuario_Cargue", "Conciliado", "Módulo Origen", "Nro. Doc. Tercero", "Núm. Fact. Pendiente de Cobrar", "Tipo Remuneración", "Referencia", "Referencia 2")
      Next n_Index
    End If
    While Not porstRecordset.EOF
      nRegistro = nRegistro + 1
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        nSecuencia = nSecuencia + 1
        poLibroExcelSave.Sheets(sHojaInterface).Range(poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1), poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33)).Font.Name = "Arial"
        poLibroExcelSave.Sheets(sHojaInterface).Range(poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1), poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33)).Font.Size = 8
        ' 1: secuencia
        sExpresion = nRegistro
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1).NumberFormat = "0"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 1).Value = sExpresion
        ' 2: estado - constante
        sExpresion = "N"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 2).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 2).Value = sExpresion
        ' 3: categoria - constante
        sExpresion = "Nomina"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 3).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 3).Value = sExpresion
        ' 4: origen - constante
        sExpresion = "INTERFACE"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 4).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 4).Value = sExpresion
        ' 5: descripción lote
        sExpresion = Trim(txtGlosa.Text)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 5).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 5).Value = sExpresion
        ' 6: nombre asiento
        sExpresion = porstRecordset!detalle
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 6).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 6).Value = sExpresion
        ' 7: descripcion asiento
        sExpresion = Trim(txtGlosa.Text)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 7).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 7).Value = sExpresion
        ' 8: compañia - constante
        sExpresion = "15"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 8).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 8).Value = sExpresion
        ' 9: cuenta contable
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codcta)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 9).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 9).Value = sExpresion
        ' 10: libre1 - constante
        sExpresion = "00"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 10).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 10).Value = sExpresion
        ' 11: area linea negocio
        sExpresion = gdl_Funcion.aTexto(porstRecordset!lineanegocio)
        sExpresion = IIf((sExpresion = "" And Left(gdl_Funcion.aTexto(porstRecordset!codcta), 1) <= "5"), "00", sExpresion)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 11).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 11).Value = sExpresion
        ' 12: concepto - seccion
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codsec)
        sExpresion = IIf((sExpresion = "" And Left(gdl_Funcion.aTexto(porstRecordset!codcta), 1) <= "5"), "00", sExpresion)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 12).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 12).Value = sExpresion
        ' 13: empresa grupo - constante
        sExpresion = "00"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 13).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 13).Value = sExpresion
        ' 14: cliente negocio
        sExpresion = gdl_Funcion.aTexto(porstRecordset!clientenego)
        sExpresion = IIf(sExpresion = "", "00000", sExpresion)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 14).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 14).Value = sExpresion
        ' 15: servicio proyecto - centro costo
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codcco)
        sExpresion = IIf(sExpresion = "", "00000", sExpresion)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 15).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 15).Value = sExpresion
        ' 16: empleado - constante
        sExpresion = "000000"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 16).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 16).Value = sExpresion
        ' 17: localizacion - segmento negocio
        sExpresion = gdl_Funcion.aTexto(porstRecordset!segmentonego)
        sExpresion = IIf(sExpresion = "", "000", sExpresion)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 17).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 17).Value = sExpresion
        ' 18: libre2 - constante
        sExpresion = "00"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 18).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 18).Value = sExpresion
        ' 19: descripción concepto
        sExpresion = gdl_Funcion.aTexto(porstRecordset!detalle)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 19).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 19).Value = sExpresion
        ' 20: debito importe moneda nacional
        n_Importe = CDec(porstRecordset!debe_mn)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 20).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 20).Value = n_Importe
        ' 21: credito importe moneda extranjera
        n_Importe = CDec(porstRecordset!haber_mn)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 21).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 21).Value = n_Importe
        ' 22: moneda
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codmon)
        sExpresion = IIf(sExpresion = s_Codmon_mn, "PEN", "USD")
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 22).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 22).Value = sExpresion
        ' 23: fecha cambio - constante
        sExpresion = ""
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 23).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 23).Value = sExpresion
        ' 24: tipo cambio - constante
        sExpresion = ""
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 24).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 24).Value = sExpresion
        ' 25: periodo liquidación - fecha yyyy/mm/dd
        sExpresion = Format(dtpFecha.Value, s_FmtFechMysql_0)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 25).NumberFormat = "yyyy/mm/dd"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 25).Value = sExpresion
        ' 26: usuario cargue - constante
        sExpresion = "0"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 26).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 26).Value = sExpresion
        ' 27: validación - constante
        sExpresion = ""
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 27).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 27).Value = sExpresion
        ' 28: modulo origen - constante
        sExpresion = ""
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 28).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 28).Value = sExpresion
        ' 29: documento tercero personal
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codpsn)
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 29).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 29).Value = sExpresion
        ' 30: docuemnto pendiente cobro - constante
        sExpresion = ""
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 30).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 30).Value = sExpresion
        ' 31: tipo remuneracion - constante
        sExpresion = ""
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 31).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 31).Value = sExpresion
        ' 32: Referencia - codigo periodo
        sExpresion = ps_Anyo & "-" & Left(cmbPeriodo.Text, 2)
        sExpresion = IIf(Left(gdl_Funcion.aTexto(porstRecordset!codcta), 1) <= "5", sExpresion, "")
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 32).NumberFormat = "@"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 32).Value = sExpresion
        ' 33: Referencia 2 - fecha proceso
        sExpresion = Format(dtpFecha.Value, s_FmtFechMysql_0)
        sExpresion = IIf(Left(gdl_Funcion.aTexto(porstRecordset!codcta), 1) <= "5", sExpresion, "")
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33).NumberFormat = "yyyy/mm/dd"
        poLibroExcelSave.Sheets(sHojaInterface).Cells(nSecuencia, 33).Value = sExpresion
      End If
      ' Incremento el porcentaje
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    If s_Accion = "G" Then
      ' Cierro objeto y saco de memoria
      sExpresion = Right(s_File, 4)
      If sExpresion = ".xls" Then
        poLibroExcelSave.SaveAs FileName:=s_File, FileFormat:=xlExcel8
      Else
        poLibroExcelSave.SaveAs FileName:=s_File, FileFormat:=xlWorkbookNormal
      End If
      poLibroExcelSave.Close SaveChanges:=False
    End If
    Set poLibroExcelSave = Nothing
    Set poAplicacionExcel = Nothing
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
Private Sub GenArchivoSap(ByVal sSufijoProceso As String, ByVal s_File As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String
  Dim sMoneda As String, sRegistro As String
  Dim sPersonal As String, sNomPersonal As String
  Dim sCencosto As String, sImpuesto As String, sDetalle As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim sCamRubro As String, s_OldMessage As String
  Dim nRegistro As Long, nRegistros As Long

  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Información ..."
  
  ' Agrupacion default
  sMoneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  ' Genero la tabla temporal de contabilización
  s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpinformacion ( "
  s_Sql = s_Sql & "codcta varchar(15) Not Null, codctax varchar(15) Null, codpsn varchar(11) Null, "
  s_Sql = s_Sql & "nombrepsn varchar(75) Null, repcodpsn char(1) Null, codcco varchar(10) Null, "
  s_Sql = s_Sql & "detalle varchar(60) Null, codmon char(1) Null, clavecon char(2) Null, "
  s_Sql = s_Sql & "tipocambio decimal(6,3) Null Default '0', fechaproceso date Null, "
  s_Sql = s_Sql & "debe_mn decimal(18,2) Null Default 0.00, haber_mn decimal(18,2) Null Default 0.00, "
  s_Sql = s_Sql & "debe_me decimal(18,2) Null Default 0.00, haber_me decimal(18,2) Null Default 0.00) "
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  If s_OptRegistro = "pllasconta" Then      ' Provisión planilla general
    ' Primer Paso : Cuentas que no tiene (centro de costo, tercero)
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & " "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & " "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Segundo Paso : Cuentas que tiene (centro de costo, tercero)
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '29' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn, dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '39' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn, dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Tercer Paso : Cuentas que tiene tercero y no centro de costo
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Cuarto Paso : Cuentas que no tiene tercero y tiene centro de costo
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Quinto Paso : Cuentas que tiene tercero y concepto; actualiza/No actualiza codigo personal
    For n_Index = 1 To 2
      s_Sql = "INSERT INTO tmpinformacion "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
      s_Sql = s_Sql & "'" & Choose(n_Index, "1", "0") & "' AS repcodpsn, Null AS codcco, cpc.descpc, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & Choose(n_Index, "2", "3") & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn, res.codcpc "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
      s_Sql = s_Sql & "'" & Choose(n_Index, "1", "0") & "' AS repcodpsn, Null AS codcco, cpc.descpc, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & Choose(n_Index, "2", "3") & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn, res.codcpc "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    Next n_Index
    
    ' Sexto Paso : Cuentas y concepto
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, cpc.descpc, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='4' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codcpc "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, cpc.descpc, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS haber_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='4' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    ' Filtrado por periodo de proceso
    If cboPeriodo.ListIndex <> 0 Then
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
    End If
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codcpc "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ElseIf s_OptRegistro = "pvscontabi" Then        ' Proviciones vaciones, gratificaciones, cts
    ' Primer Paso : Cuentas que no tiene (centro de costo, tercero)
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & " "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & " "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Segundo Paso : Cuentas que tiene (centro de costo, tercero)
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '29' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=pdo.codcls AND dxc.codpdo=pdo.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn, dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '39' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=pdo.codcls AND dxc.codpdo=pdo.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn, dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Tercer Paso : Cuentas que tiene tercero y no centro de costo
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & "'" & s_Estado_Act & "' AS repcodpsn, Null AS codcco, Null AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Cuarto Paso : Cuentas que no tiene tercero y tiene centro de costo
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=pdo.codcls AND dxc.codpdo=pdo.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, dxc.codcco, Null AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=pdo.codcls AND dxc.codpdo=pdo.codpdo AND dxc.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", dxc.codcco "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    
    ' Quinto Paso : Cuentas que tiene tercero y concepto; actualiza/No actualiza codigo personal
    For n_Index = 1 To 2
      s_Sql = "INSERT INTO tmpinformacion "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
      s_Sql = s_Sql & "'" & Choose(n_Index, "1", "0") & "' AS repcodpsn, Null AS codcco, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & Trim(cmbProceso.Text) & "'") & " AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & Choose(n_Index, "2", "3") & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn" & IIf(cmbProceso.ListIndex = 2, ", res.codcpc ", " ")
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
      s_Sql = s_Sql & "'" & Choose(n_Index, "1", "0") & "' AS repcodpsn, Null AS codcco, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & Trim(cmbProceso.Text) & "'") & " AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & Choose(n_Index, "2", "3") & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn" & IIf(cmbProceso.ListIndex = 2, ", res.codcpc", "") & " "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
    Next n_Index
    
    ' Sexto Paso : Cuentas y concepto
    s_Sql = "INSERT INTO tmpinformacion "
    s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & Trim(cmbProceso.Text) & "'") & " AS detalle, res.codmon, '40' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='4' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & IIf(cmbProceso.ListIndex = 2, ", res.codcpc ", " ")
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, cta.codcta_ajd_hab AS codctax, Null AS codpsn, Null AS nombrepsn, "
    s_Sql = s_Sql & "Null AS repcodpsn, Null AS codcco, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & Trim(cmbProceso.Text) & "'") & " AS detalle, res.codmon, '50' AS clavecon, pdo.tipocambio, pdo.fechaproceso, "
    s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS haber_mn, "
    s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS haber_me "
    If cmbProceso.ListIndex = 2 Then
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    Else
      s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
    End If
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "AND cta.inddoc='4' AND cta.indcco='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
    s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & IIf(cmbProceso.ListIndex = 2, ", res.codcpc", "") & " "
    s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  End If
  
  ' Ultimo-1 Paso : Actualizo los codigos de deudor
  s_Sql = "UPDATE tmpinformacion det, plpersonal psn "
  s_Sql = s_Sql & "SET det.codcta=psn.coddeudor "
  s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND det.codpsn=psn.codpsn "
  s_Sql = s_Sql & "AND IFNULL(det.repcodpsn, '')<>'" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND IFNULL(psn.coddeudor, '')<>'' "
  s_Sql = s_Sql & "AND IFNULL(det.debe_mn, 0)<>0"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar
  ' Ultimo-2 Paso : Actualizo los codigos de acreedor
  s_Sql = "UPDATE tmpinformacion det, plpersonal psn "
  s_Sql = s_Sql & "SET det.codcta=psn.codacredor "
  s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND det.codpsn=psn.codpsn "
  s_Sql = s_Sql & "AND IFNULL(det.repcodpsn, '')<>'" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND IFNULL(psn.codacredor, '')<>'' "
  s_Sql = s_Sql & "AND IFNULL(det.haber_mn, 0)<>0"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar
  
  ' Registros de contabilización
  s_Sql = "SELECT tmp.codcta, tmp.codctax, tmp.codpsn, tmp.nombrepsn, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "tmp.clavecon, tmp.tipocambio, tmp.fechaproceso, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me, IFNULL(tmp.detalle, cta.detcta) AS detalle "
  s_Sql = s_Sql & "FROM tmpinformacion tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    ' Muestro la Barra
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    sMoneda = IIf(sMoneda = "mn", "PEN", "USD")
    nTipoCambio = CDec(porstRecordset!Tipocambio)
    
    ' Creo objeto de archivo
    Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
    Set potxtFileExp = pofsoFileExp.CreateTextFile(s_File, True)
    s_Caracter = "/"
    ' Genero cabecera del archivo
    psRegistro = "1FB01" & Space(16) & s_Caracter
    psRegistro = psRegistro & Space(29) & Format(porstRecordset!fechaproceso, "ddmmyyyy")
    psRegistro = psRegistro & "ZNPE10" & Format(porstRecordset!fechaproceso, "ddmmyyyy") & sMoneda & Space(2)
    sRegistro = Left(Trim(txtGlosa.Text), 16)
    psRegistro = psRegistro & gdl_Funcion.PadR(sRegistro, 16, " ")
    psRegistro = psRegistro & gdl_Funcion.PadR(Left(sRegistro, 15), 15, " ")
    sRegistro = Left(Trim(cmbPeriodo.Text), 2) & ps_Anyo
    psRegistro = psRegistro & gdl_Funcion.PadR(sRegistro, 6, " ")
    psRegistro = psRegistro & "0100" & Format(porstRecordset!fechaproceso, "ddmmyyyy") & s_Caracter & " " & s_Caracter
    psRegistro = psRegistro & Space(9) & s_Caracter & s_Caracter & Space(15) & s_Caracter & Space(3) & s_Caracter
    psRegistro = psRegistro & Space(7) & s_Caracter & s_Caracter & Space(17) & s_Caracter & Space(49) & s_Caracter
    psRegistro = psRegistro & Space(9) & s_Caracter & Space(9) & s_Caracter
    psRegistro = psRegistro & gdl_Funcion.PadL(Format(nTipoCambio, "##0.###0"), 17, " ") & Space(4)
    psRegistro = psRegistro & s_Caracter & s_Caracter & " " & s_Caracter & " " & s_Caracter & Space(13) & s_Caracter & " " & s_Caracter & s_Caracter & Space(7) & s_Caracter
    potxtFileExp.WriteLine psRegistro
    
    ' Genero el detalle
    While Not porstRecordset.EOF
      psRegistro = "2" & s_Caracter & Space(19)
      psRegistro = psRegistro & "BBSEG" & Space(25) & s_Caracter & Space(7) & s_Caracter
      psRegistro = psRegistro & " " & s_Caracter & Space(3) & s_Caracter & Space(7) & s_Caracter
      psRegistro = psRegistro & Space(4) & s_Caracter & Space(15) & s_Caracter & Space(24) & s_Caracter & Space(7)
      sRegistro = Left(gdl_Funcion.aTexto(porstRecordset!codctax), 3)
      psRegistro = psRegistro & Trim(IIf((sRegistro = "141" Or sRegistro = "385" Or sRegistro = "389"), "39", porstRecordset!clavecon))
      psRegistro = psRegistro & gdl_Funcion.PadR(porstRecordset!codcta, 10, " ")
      sRegistro = gdl_Funcion.aTexto(porstRecordset!codctax)
      psRegistro = psRegistro & IIf(Left(sRegistro, 3) = "141", "R", IIf(Left(sRegistro, 3) = "385", "A", IIf(Left(sRegistro, 3) = "389", "E", s_Caracter)))
      nImporte = CDec(porstRecordset!debe_mn) + CDec(porstRecordset!haber_mn)
      psRegistro = psRegistro & gdl_Funcion.PadL(Format(nImporte, "############0.00"), 16, " ") & s_Caracter & Space(3)
      psRegistro = psRegistro & Format(porstRecordset!fechaproceso, "ddmmyyyy") & s_Caracter
      psRegistro = psRegistro & gdl_Funcion.PadR(IIf((Left(sRegistro, 3) = "141" Or Left(sRegistro, 3) = "385" Or Left(sRegistro, 3) = "389"), porstRecordset!codcta, Format(dtpFecha.Value, "mm/yy")), 10, " ") & Space(8)
      psRegistro = psRegistro & gdl_Funcion.PadR(sRegistro, 8, " ")
      psRegistro = psRegistro & gdl_Funcion.PadR(porstRecordset!detalle, 42, " ")
      sRegistro = gdl_Funcion.aTexto(porstRecordset!codcco)
      psRegistro = psRegistro & gdl_Funcion.PadR(IIf(sRegistro = "", s_Caracter, sRegistro), 10, " ") & s_Caracter
      psRegistro = psRegistro & Space(9) & s_Caracter & Space(11) & s_Caracter & Space(9) & s_Caracter & s_Caracter
      psRegistro = psRegistro & " " & s_Caracter & " " & s_Caracter & Space(13) & s_Caracter & " " & s_Caracter & s_Caracter
      nImporte = CDec(porstRecordset!debe_me) + CDec(porstRecordset!haber_me)
      psRegistro = psRegistro & gdl_Funcion.PadL(Format(nImporte, "###########0.00"), 23, " ")
      potxtFileExp.WriteLine psRegistro
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      porstRecordset.MoveNext
    Wend
    ' Cierro objeto y saco de memoria
    potxtFileExp.Close
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje s_OldMessage
  ' Elimino la tabla temporal de contabilización
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, "DROP TABLE IF EXISTS tmpinformacion"
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub GenArchivoSapExcel(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim poApplExcel As Object, poLibroExcel As Object
  Dim sHojaExcel As String, sMoneda As String
  Dim sExpresion As String, s_OldMessage As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nSecuencia As Long

  ' Genero la tabla con información
  RecuperaRegistros s_Archivo
  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me, "
  s_Sql = s_Sql & "IFNULL(cco.detcco,cta.detcta) AS detalleitem "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco AND cco.estcco='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set poApplExcel = CreateObject("Excel.Application")
      poApplExcel.Visible = False
      sExpresion = Trim(cmbPeriodo.Text)
      Set poLibroExcel = poApplExcel.Workbooks.Add
      sHojaExcel = Left(sExpresion, 20)
      poLibroExcel.Sheets("Hoja1").Name = sHojaExcel
    
      nSecuencia = 1
      ' Titulos de registro
      sExpresion = "Serial"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
      
      sExpresion = "Document type"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
      
      sExpresion = "Posting date"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
      
      sExpresion = "Document date"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
      
      sExpresion = "Currency"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
      
      sExpresion = "Reference document"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
      
      sExpresion = "Document Header"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
      
      sExpresion = "Posting key"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
      
      sExpresion = "Special G/L"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
      
      sExpresion = "G/L account"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
      
      sExpresion = "Vendor code"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
      
      sExpresion = "Customer code"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).Value = sExpresion
      
      sExpresion = "Main asset"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).Value = sExpresion
      
      sExpresion = "Sub asset"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).Value = sExpresion
      
      sExpresion = "Asset type"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).Value = sExpresion
      
      sExpresion = "Base amount"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).Value = sExpresion
      
      sExpresion = "Tax amount"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).Value = sExpresion
      
      sExpresion = "Document currency"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).Value = sExpresion
      
      sExpresion = "Local currency"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).Value = sExpresion
      
      sExpresion = "Business area"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).Value = sExpresion
      
      sExpresion = "Profit center"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).Value = sExpresion
      
      sExpresion = "Cost center"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).Value = sExpresion
      
      sExpresion = "WBS"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).Value = sExpresion
      
      sExpresion = "Tax code"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).Value = sExpresion
      
      sExpresion = "Business place"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).Value = sExpresion
      
      sExpresion = "Assignment number"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).Value = sExpresion
      
      sExpresion = "XREF1"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).Value = sExpresion
      
      sExpresion = "XREF2"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).Value = sExpresion
      
      sExpresion = "XREF3"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).Value = sExpresion
      
      sExpresion = "Payment method"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).Value = sExpresion
      
      sExpresion = "Payment term"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).Value = sExpresion
      
      sExpresion = "Payment block"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).Value = sExpresion
      
      sExpresion = "Partner bank type"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).Value = sExpresion
      
      sExpresion = "Due date"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).Value = sExpresion
      
      sExpresion = "Item text"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).Value = sExpresion
      
      sExpresion = "WBS2"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).Value = sExpresion
      
      sExpresion = "Profit cen"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 37).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 37).Value = sExpresion
    End If
    
    nSecuencia = 2
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        ' detalle por moneda
        sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
        If porstRecordset!codmon = sMoneda Then
         ' Serial
          sExpresion = Trim(txtComprobante.Text)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
          ' Document type - constante
          sExpresion = "PS"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
          ' Posting date
          sExpresion = Format(dtpFecha.Value, "yyyymmdd")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
          ' Document date
          sExpresion = Format(dtpFecha.Value, "yyyymmdd")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
          ' Currency
          sMoneda = IIf(porstRecordset!codmon = s_Codmon_mn, "PEN", "USD")
          sExpresion = sMoneda
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
          ' Reference document
          sExpresion = "PR" & Trim(txtDiario.Text) & Left(Trim(cmbPeriodo.Text), 2) & ps_Anyo
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
          ' Document Header
          sExpresion = Trim(txtGlosa.Text)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
          ' Posting key
          sExpresion = IIf(CDec(porstRecordset("debe_m" & porstRecordset!codmon)) > 0, "40", "50")
          If porstRecordset!inddoc = s_Estado_Act Then
            sExpresion = IIf(CDec(porstRecordset("debe_m" & porstRecordset!codmon)) > 0, "21", "31")
          End If
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
          ' Special G/L - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
          ' G/L account
          sExpresion = gdl_Funcion.aTexto(porstRecordset!codcta)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
          ' Vendor code
          sExpresion = gdl_Funcion.aTexto(porstRecordset!codpsn)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
          ' Customer code - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).Value = sExpresion
          ' Main asset - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).Value = sExpresion
          ' Sub asset - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).Value = sExpresion
          ' Asset type - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).Value = sExpresion
          ' Base amount - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).Value = sExpresion
          ' Tax amount - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).Value = sExpresion
          ' Document currency
          nImporte = CDec(porstRecordset("debe_m" & porstRecordset!codmon)) + CDec(porstRecordset("haber_m" & porstRecordset!codmon))
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).NumberFormat = "#,##0.00"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).Value = nImporte
          ' Local currency - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).Value = sExpresion
          ' Business area - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).Value = sExpresion
          ' Profit center - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).Value = sExpresion
          ' Cost center
          sExpresion = gdl_Funcion.aTexto(porstRecordset!codcco)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).Value = sExpresion
          ' WBS - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).Value = sExpresion
          ' Tax code - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).Value = sExpresion
          ' Business place - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 25).Value = sExpresion
          ' Assignment number  - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 26).Value = sExpresion
          ' XREF1 . constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 27).Value = sExpresion
          ' XREF2 - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 28).Value = sExpresion
          ' XREF3 - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 29).Value = sExpresion
          ' Payment method - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 30).Value = sExpresion
          ' Payment term
          sExpresion = IIf(Left(porstRecordset!codcta, 1) = "4", "0001", "")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 31).Value = sExpresion
          ' Payment block - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 32).Value = sExpresion
          ' Partner bank type - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 33).Value = sExpresion
          ' Due date
          sExpresion = Format(dtpFecha.Value, "yyyymmdd")
          sExpresion = IIf(Left(porstRecordset!codcta, 1) = "4", sExpresion, "")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 34).Value = sExpresion
          ' Item text
          sExpresion = gdl_Funcion.aTexto(porstRecordset!detalleitem)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 35).Value = sExpresion
          ' WBS2 - constante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 36).Value = sExpresion
          ' Profit cen - cosntante
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 37).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 37).Value = sExpresion
        End If
      End If
      ' Incremento el porcentaje
      nSecuencia = nSecuencia + 1
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend

    If s_Accion = "G" Then
      ' Cierro y grabo documento excel
      sExpresion = Strings.Right(s_File, 4)
      If sExpresion = ".xls" Then
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlExcel8
      Else
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlWorkbookNormal
      End If
      poLibroExcel.Close SaveChanges:=False
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Saco de memoria objeto
  Set poLibroExcel = Nothing
  Set poApplExcel = Nothing
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub

Private Sub GenArchivoSpring(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String
  Dim n_Importe As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim sExpresion As String, s_OldMessage As String
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo

  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_File, True)
      s_Caracter = Chr(vbKeyTab)
    End If
    While Not porstRecordset.EOF
      nRegistro = nRegistro + 1
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = ""
        ' 1: cuenta contable
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codcta) & s_Caracter
        ' 2: persona
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codpsn) & s_Caracter
        ' 3: proyecto
        sExpresion = Trim(txtDiario.Text) & IIf(Left(Trim(txtDiario.Text), 1) = "Z", "-00-00", "-99-99")
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 4: fecha
        psRegistro = psRegistro & Format(dtpFecha, s_FormatoFecha) & s_Caracter
        ' 5: centro de costos
        psRegistro = psRegistro & Mid(gdl_Funcion.aTexto(porstRecordset!codcco), 2) & s_Caracter
        ' 6: documento
        sExpresion = "PLLA" & "-" & ps_Anyo & Left(Trim(cmbPeriodo.Text), 2)
        sExpresion = IIf((gdl_Funcion.aTexto(porstRecordset!inddoc) = s_Estado_Act Or gdl_Funcion.aTexto(porstRecordset!codcta) = "40170004"), sExpresion, "")
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 7: sucursal - constante
        sExpresion = "LIMA"
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 8: referencia - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 9: importe moneda nacional
        n_Importe = CDec(porstRecordset!debe_mn) + CDec(porstRecordset!haber_mn)
        n_Importe = n_Importe * IIf(porstRecordset!debe_mn > 0, 1, -1)
        psRegistro = psRegistro & Format(n_Importe, "###########0.00") & s_Caracter
        ' 10: importe moneda extranjera
        n_Importe = CDec(porstRecordset!debe_me) + CDec(porstRecordset!haber_me)
        n_Importe = n_Importe * IIf(porstRecordset!debe_mn > 0, 1, -1)
        psRegistro = psRegistro & Format(n_Importe, "###########0.00") & s_Caracter
        ' 11: referencia1 - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 12: referencia9 - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 13: descripcion o glosa
        psRegistro = psRegistro & Left(Trim(txtGlosa.Text), 40) & s_Caracter
        ' 14: centro costo destino - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 15: inter compañia - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 16: reparo - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 17: referenciax - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        ' 18: orden compra - constante
        sExpresion = ""
        psRegistro = psRegistro & sExpresion & s_Caracter
        potxtFileExp.WriteLine psRegistro
      End If
      ' Incremento el porcentaje
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    If s_Accion = "G" Then
      ' Cierro objeto y saco de memoria
      potxtFileExp.Close
    End If
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
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
Private Sub GenArchivoSpring_Proyecto(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String
  Dim n_Importe As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim sExpresion As String, s_OldMessage As String
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo

  nTipoCambio = 1
  ' Obtengo el tipo de cambio
  s_Sql = "SELECT codpdo, tipocambio "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
  End If
  porstRecordset.Close

  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me, "
  s_Sql = s_Sql & "sec.codintersec, ubi.codinterubica "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "LEFT JOIN plseccion sec ON sec.codsec=tmp.codsec "
  s_Sql = s_Sql & "LEFT JOIN plubicacion ubi ON ubi.codubica=tmp.codubica "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_File, True)
      s_Caracter = " "
      
      ' cabecera
      psRegistro = ""
      ' 1: tipo registro - constante
      sExpresion = "C"
      psRegistro = psRegistro & sExpresion
      ' 2: periodo
      sExpresion = ps_Anyo & Left(cmbPeriodo.Text, 2)
      psRegistro = psRegistro & sExpresion
      ' 3: compañia socio - constante
      sExpresion = "00002500"
      psRegistro = psRegistro & sExpresion
      ' 4: numero voucher
      sExpresion = Trim(txtComprobante.Text)
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 6, s_Caracter)
      ' 5: numeracion automatica - constante
      sExpresion = "S"
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 1, s_Caracter)
      ' 6: unidad de negocio - constante
      sExpresion = "SECE"
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 4, s_Caracter)
      ' 7: moneda
      sExpresion = gdl_Funcion.aTexto(porstRecordset!codmon)
      sExpresion = IIf(sExpresion = s_Codmon_mn, "LO", "EX")
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 2, s_Caracter)
      ' 8: lote vouchers - constante
      sExpresion = ""
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 6, s_Caracter)
      ' 9: descripcion o glosa
      sExpresion = Trim(txtGlosa.Text)
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 50, s_Caracter)
      ' 10: fecha voucher
      sExpresion = Format(dtpFecha.Value, "yyyymmdd")
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 8, s_Caracter)
      ' 11: importe moneda nacional
      n_Importe = CDec(nTipoCambio)
      psRegistro = psRegistro & gdl_Funcion.PadL(Format(n_Importe, "#####0.000"), 10, "0")
      ' 12: numero interno
      sExpresion = "PLLA" & ps_Anyo & Left(Trim(cmbPeriodo.Text), 2)
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 10, s_Caracter)
      ' 13: libro contable - constante
      sExpresion = ""
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 2, s_Caracter)
      ' 14: clasificacion - constante
      sExpresion = ""
      psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 4, s_Caracter)
      potxtFileExp.WriteLine psRegistro
    End If
    While Not porstRecordset.EOF
      nRegistro = nRegistro + 1
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = ""
        
        ' 1: tipo registro - constante
        sExpresion = "D"
        psRegistro = psRegistro & sExpresion
        ' 2: periodo
        sExpresion = ps_Anyo & Left(cmbPeriodo.Text, 2)
        psRegistro = psRegistro & sExpresion
        ' 3: compañia socio - constante
        sExpresion = "00002500"
        psRegistro = psRegistro & sExpresion
        ' 4: numero voucher
        sExpresion = Trim(txtComprobante.Text)
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 6, s_Caracter)
        ' 5: secuencia voucher
        sExpresion = nRegistro
        psRegistro = psRegistro & gdl_Funcion.PadL(sExpresion, 4, "0")
        ' 6: cuenta contable
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codcta)
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 20, s_Caracter)
        ' 7: indicador cuenta equivalente - constante
        sExpresion = "N"
        psRegistro = psRegistro & sExpresion
        ' 8: centro de costos
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codcco)
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 10, s_Caracter)
        ' 9: destino centro de costos - constante
        sExpresion = ""
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 6, s_Caracter)
        ' 10: proyecto
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codinterubica)
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 15, s_Caracter)
        ' 11: codigo persona
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codpsn)
        psRegistro = psRegistro & gdl_Funcion.PadL(sExpresion, 6, s_Caracter)
        ' 12: importe moneda nacional
        n_Importe = CDec(porstRecordset!debe_mn) + CDec(porstRecordset!haber_mn)
        sExpresion = gdl_Funcion.PadL(Format(n_Importe, "###########0.00"), 14, "0")
        sExpresion = IIf(porstRecordset!debe_mn > 0, "0", "-") & sExpresion
        psRegistro = psRegistro & sExpresion
        ' 13: importe moneda extranjera
        n_Importe = CDec(porstRecordset!debe_me) + CDec(porstRecordset!haber_me)
        sExpresion = gdl_Funcion.PadL(Format(n_Importe, "###########0.00"), 14, "0")
        sExpresion = IIf(porstRecordset!debe_mn > 0, "0", "-") & sExpresion
        psRegistro = psRegistro & sExpresion
        ' 14: fecha
        sExpresion = Format(dtpFecha.Value, "yyyymmdd")
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 8, s_Caracter)
        ' 15: documento referencia - constante
        sExpresion = ""
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 20, s_Caracter)
        ' 16: sucursal
        sExpresion = gdl_Funcion.aTexto(porstRecordset!codintersec)
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 4, s_Caracter)
        ' 17: descripcion o glosa
        sExpresion = Trim(txtGlosa.Text)
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 50, s_Caracter)
        ' 18: persona spring - constante
        sExpresion = ""
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 20, s_Caracter)
        ' 19: referencia - constante
        sExpresion = ""
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 12, s_Caracter)
        ' 20: contrato - constante
        sExpresion = ""
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 10, s_Caracter)
        ' 21: inter compañia - constante
        sExpresion = ""
        psRegistro = psRegistro & gdl_Funcion.PadR(sExpresion, 8, s_Caracter)
        potxtFileExp.WriteLine psRegistro
      End If
      ' Incremento el porcentaje
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    If s_Accion = "G" Then
      ' Cierro objeto y saco de memoria
      potxtFileExp.Close
    End If
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
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
Public Sub GenArchivoInformatExcel(ByVal s_Archivo As String, ByVal s_File As String, ByVal s_Accion As String)
  Dim poApplExcel As Object, poLibroExcel As Object
  Dim sHojaExcel As String, sMoneda As String
  Dim sExpresion As String, s_OldMessage As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nSecuencia As Long
  Dim dFechaTCambio As Date
  
 
  ' Inicializando variable status validaciones
  s_StatusValid_DatosConcar = "OK"
  ' Genero la tabla con información
  RecuperaRegistros s_Archivo
  
  nTipoCambio = 1
  dFechaTCambio = Format(dtpFecha, s_FmtFechMysql_0)
  ' Obtengo el tipo de cambio
  s_Sql = "SELECT codpdo, tipocambio,fechaproceso "
  s_Sql = s_Sql & "FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
  ' Filtrado por periodo de proceso
  If cboPeriodo.ListIndex <> 0 Then
    s_Sql = s_Sql & "AND codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
  End If
  s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpdo DESC"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    nTipoCambio = CDec(porstRecordset!Tipocambio)
    dFechaTCambio = porstRecordset!fechaproceso
  End If
  porstRecordset.Close
  
  ' Recupero la información para exportar
  s_Sql = "SELECT tmp.codcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.detalle, tmp.codmon, "
  s_Sql = s_Sql & "cta.tpotcb, cta.inddoc, tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me, "
  s_Sql = s_Sql & "IFNULL(cco.detcco,cta.detcta) AS detalleitem, "
  s_Sql = s_Sql & "tmp.codsec, tmp.codubica, ubi.codinterubica "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "LEFT JOIN plpersonal per ON tmp.codpsn = per.codpsn "
  's_Sql = s_Sql & "LEFT JOIN plcencospro cos ON tmp.codpsn = cos.codpsn "
  s_Sql = s_Sql & "LEFT JOIN plubicacion ubi ON ubi.codubica=tmp.codubica "
  s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco AND cco.estcco='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "ORDER BY codcta"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0

    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("diario", "comprobante", "fecha", "glosa", "codcta", "codpsn", "codcco", "detalle", "codmon", "tipcambio", "debe_mn", "haber_mn", "debe_me", "haber_me")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set poApplExcel = CreateObject("Excel.Application")
      poApplExcel.Visible = False
      sExpresion = Trim(cmbPeriodo.Text)
      Set poLibroExcel = poApplExcel.Workbooks.Add
      sHojaExcel = Left(sExpresion, 20)
      poLibroExcel.Sheets("Hoja1").Name = sHojaExcel
      
      'nSecuencia = 1
      nSecuencia = 10
      ' Titulos de registro
      sExpresion = "Fecha de Registro"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
      sExpresion = "Tipo de Documento NAV 0 - Vacio 1 - Pago 2 - Factura 3 - Abono"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = sExpresion
      sExpresion = "Codigo Tipo documento, segun SUNAT (ver Pestaña TipoDocumento) "
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
      
      sExpresion = "N° Documento (max. 20 caracteres)"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
     
      sExpresion = "Nº Documento Externo Adicional (max. 20 caracteres)"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
      sExpresion = "Grupo Contable (Según pestaña Grupos contables proveedor)"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
      sExpresion = "Fecha Emision Documento"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
      sExpresion = "Fecha Vencimiento"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
       sExpresion = "Tipo mov. (Fijo)"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
      sExpresion = "Nº Codigo NAV/ Ruc - Provedor /DNI Trabajador"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
      sExpresion = "Descripcion (max. 50 caracteres)"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
      sExpresion = "Importe"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).Value = sExpresion
      sExpresion = "Moneda (" & s_Codmon_mn_Nom & " en Blanco)"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).Value = sExpresion
      sExpresion = "Tipo de Cambio"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).Value = sExpresion
      sExpresion = "Cod Comprador"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).Value = sExpresion
      sExpresion = "Tipo Destino"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).Value = sExpresion
      sExpresion = "Producto"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).Value = sExpresion
      sExpresion = "Categoria"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).Value = sExpresion
      sExpresion = "Sucursal"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).Value = sExpresion
      sExpresion = "Tamaño"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).Value = sExpresion
      sExpresion = "Dimension 5"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).Value = sExpresion
      sExpresion = "Dimension 6"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).Value = sExpresion
      sExpresion = "Dimension 7"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).Value = sExpresion
      sExpresion = "Dimension 8"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).NumberFormat = "@"
      poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).Value = sExpresion
      
      
    End If
    'nSecuencia = 2
    nSecuencia = 11
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(Trim(txtDiario.Text), Trim(txtComprobante.Text), Format(dtpFecha, s_FmtFechMysql_0), Trim(txtGlosa.Text), gdl_Funcion.aTexto(porstRecordset("codcta")), gdl_Funcion.aTexto(porstRecordset("codpsn")), gdl_Funcion.aTexto(porstRecordset("codcco")), gdl_Funcion.aTexto(porstRecordset("detalle")), gdl_Funcion.aTexto(porstRecordset("codmon")), CDec(nTipoCambio), CDec(porstRecordset("debe_mn")), CDec(porstRecordset("haber_mn")), CDec(porstRecordset("debe_me")), CDec(porstRecordset("haber_me")))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_File, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        ' detalle por moneda
        sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
        If porstRecordset!codmon = sMoneda Then
        
         ' Fecha de comprobante
          sExpresion = Format(dtpFecha.Value, "dd/mm/yyyy")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 1).Value = sExpresion
         ' Tipo de Documento
         ' nImporte = 0
         ' poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
          
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 2).Value = ""
          ' Codigo Tipo de Documento
          'sExpresion = "01"
          'poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
          'poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = sExpresion
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 3).Value = ""
          
          'Numero de Documento
          sExpresion = txtGlosa.Text
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 4).Value = sExpresion
          'Nª Documento Externo Adicional
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 5).Value = sExpresion
          'Nª Grupo Contable
          sExpresion = porstRecordset!codcta
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 6).Value = sExpresion
          ' Fecha de emision del documento
          sExpresion = Format(dtpFecha.Value, "dd/mm/yyyy")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 7).Value = sExpresion
'          ' Fecha de Vencimiento
          sExpresion = Format(dtpFecha.Value, "dd/mm/yyyy")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 8).Value = sExpresion
          
          'Tipo de Movimiento
          If porstRecordset!inddoc = 0 Then
            
          
             If Left(IIf(IsNull(porstRecordset!codcta), " ", porstRecordset!codcta), 2) = "42" Then
                sExpresion = "2"
             Else
                sExpresion = "0"
             End If
          Else
             sExpresion = "1"  '{0=AFECTA CUENTA CONTABLE,1 = AFECTA CLIENTE,2=AFECTA PROVEEDOR}
          End If
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 9).Value = sExpresion
          
          'Nª de Codigo
          'sExpresion = porstRecordset!codcta
           If porstRecordset!inddoc = 0 Then
           sExpresion = IIf(IsNull(porstRecordset!codcta), " ", porstRecordset!codcta)
           Else
            sExpresion = IIf(IsNull(porstRecordset!codpsn), " ", porstRecordset!codpsn)
           End If
          
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 10).Value = sExpresion
          'Glosa Principal
          sExpresion = porstRecordset!detalle
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 11).Value = sExpresion
          'Importe
          sExpresion = IIf(porstRecordset!debe_mn > 0, "D", "H")
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).NumberFormat = "#,##0.00"
          If sExpresion = "D" Then
            nImporte = (CDec(porstRecordset("debe_m" & porstRecordset!codmon)))
          ElseIf sExpresion = "H" Then
            nImporte = CDec(porstRecordset("haber_m" & porstRecordset!codmon)) * -1
          End If
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 12).Value = nImporte
          
          ' Codigo de Moneda
          sMoneda = IIf(porstRecordset!codmon = s_Codmon_mn, "MN", "US")
          sExpresion = sMoneda
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).NumberFormat = "@"
          If sExpresion = "MN" Then
             sExpresion = ""
          End If
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 13).Value = sExpresion
          ' Tipo de Cambio
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 14).Value = sExpresion
          'Cod Comprador
          sExpresion = ""
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 15).Value = sExpresion
          'Tipo Destino , PREGUNTAR
          sExpresion = IIf(IsNull(porstRecordset!codsec) = True, "", porstRecordset!codsec)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 16).Value = sExpresion
          'Producto
          If Left(porstRecordset!codcta, 1) = 6 Then 'Las Cuentas 6 son de Gasto
          sExpresion = "NA"
          Else
          sExpresion = ""
          End If
          
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 17).Value = sExpresion
          'Categoria
          sExpresion = "EB001"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 18).Value = sExpresion
          'Sucursal, PREGUNTAR
          sExpresion = IIf(IsNull(porstRecordset!codinterubica) = True, "", porstRecordset!codinterubica)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 19).Value = sExpresion
          'Tamaño
          If Left(porstRecordset!codcta, 1) = 6 Then 'Las Cuentas 6 son de Gasto
          sExpresion = "TA003"
          Else
          sExpresion = ""
          End If
          
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 20).Value = sExpresion
          'Dimension 5, PREGUNTAR
          sExpresion = IIf(IsNull(porstRecordset!codcco) = True, "", porstRecordset!codcco)
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).NumberFormat = "@"
          poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 21).Value = sExpresion
          'Dimension 6
           If Left(porstRecordset!codcta, 1) = 6 Then 'Las Cuentas 6 son de Gasto
           sExpresion = "NA"
           Else
           sExpresion = ""
           End If
           
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).NumberFormat = "@"
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 22).Value = sExpresion
           'Dimension 7
           sExpresion = ""
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).NumberFormat = "@"
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 23).Value = sExpresion
           'Dimension 8
           sExpresion = ""
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).NumberFormat = "@"
           poLibroExcel.Sheets(sHojaExcel).Cells(nSecuencia, 24).Value = sExpresion
          
        End If
      End If
      ' Incremento el porcentaje
      nSecuencia = nSecuencia + 1
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    
    'Formateando la Altura de la Cabecera de Rotulo
       poLibroExcel.Sheets(sHojaExcel).Rows("10:10").RowHeight = 89
       poLibroExcel.Sheets(sHojaExcel).Range("A10:X10").EntireColumn.AutoFit
       'poLibroExcel.Sheets(sHojaExcel).Range("A10:X10").EntireRow.AutoFit
       
       poLibroExcel.Sheets(sHojaExcel).Range("A10:X10").Select
       
       With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
       End With
        
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    If s_Accion = "G" Then
      ' Cierro y grabo documento excel
      sExpresion = Strings.Right(s_File, 4)
      If sExpresion = ".xls" Then
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlExcel8
      Else
        poLibroExcel.SaveAs FileName:=s_File, FileFormat:=xlOpenXMLWorkbook
         
      End If
      poLibroExcel.Close SaveChanges:=False
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Saco de memoria objeto
  Set poLibroExcel = Nothing
  Set poApplExcel = Nothing
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  

End Sub
Private Sub RecuperaRegistros(ByVal sArchivo As String)
  Dim sMoneda As String, sSufijoProceso As String
  
  sSufijoProceso = IIf(chkProceso.Value, "_pdo", "")
  For n_Index = 1 To 2
    sMoneda = Choose(n_Index, "mn", "me")
    If s_OptRegistro = "pllasconta" Then
      ' Primer Paso : Cuentas que no tiene (centro de costo, tercero)
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, Null, Null, Null, Null, Null, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codcpc "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, Null, Null, Null, Null, Null, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codcpc "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
      
      ' Segundo Paso : Cuentas que tiene (centro de costo, tercero)
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codcpc, psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codcpc, psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
      
      ' Tercer Paso : Cuentas que tiene tercero y no centro de costo
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, Null, Null, Null, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codcpc, psn.codacredor, res.codpsn "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, Null, Null, Null, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codcpc, psn.codacredor, res.codpsn "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
      
      ' Cuarto Paso : Cuentas que no tiene tercero y tiene centro de costo
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, Null, Null, dxc.codcco, dxr.codsec, dxr.codubica, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codcpc, dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, Null, Null, dxc.codcco, dxr.codsec, dxr.codubica, cpc.descpc, res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codproce" & sSufijoProceso & "='" & Right(Trim(cmbProceso.Text), 2) & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' Filtrado por periodo de proceso
      If cboPeriodo.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND res.codpdo='" & Trim(Left(cboPeriodo.Text, 8)) & "' "
      End If
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codcpc, dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
    ElseIf s_OptRegistro = "pvscontabi" Then
      ' Primer Paso : Cuentas que no tiene (centro de costo, tercero)
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, Null, Null, Null, Null, Null, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & " "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, Null, Null, Null, Null, Null, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & " "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
        
      ' Segundo Paso : Cuentas que tiene (centro de costo, tercero)
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON pdo.codcls=dxc.codcls AND pdo.codpdo=dxc.codpdo AND res.codpsn=dxc.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, dxc.codcco, dxr.codsec, dxr.codubica, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON pdo.codcls=dxc.codcls AND pdo.codpdo=dxc.codpdo AND res.codpsn=dxc.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn, psn.codacredor, dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
      
      ' Tercer Paso : Cuentas que tiene tercero y no centro de costo
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, Null, Null, Null, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS debe_me, 0.00 AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn, psn.codacredor "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, psn.codacredor, res.codpsn, Null, Null, Null, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn, 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me, 0)), 2) AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", psn.codacredor, res.codpsn "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
      
      ' Cuarto Paso : Cuentas que no tiene tercero y tiene centro de costo
      s_Sql = "INSERT INTO " & sArchivo & " "
      s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, Null, Null, dxc.codcco, Null, Null, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS debe_mn, 0.00 AS haber_mn, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS debe_me, 0.00 AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON pdo.codcls=dxc.codcls AND pdo.codpdo=dxc.codpdo AND res.codpsn=dxc.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00) "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, Null, Null, dxc.codcco, Null, Null, " & IIf(cmbProceso.ListIndex = 2, "cpc.descpc", "'" & UCase(Trim(cmbProceso.Text)) & "'") & ", res.codmon, "
      s_Sql = s_Sql & "0.00 AS debe_mn, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_mn*(dxc.porcentaje/100), 0)), 2) AS haber_mn, "
      s_Sql = s_Sql & "0.00 AS debe_me, ROUND(SUM(IFNULL(res.importe" & IIf(cmbProceso.ListIndex = 2, "", "pvs") & "_me*(dxc.porcentaje/100), 0)), 2) AS haber_me "
      If cmbProceso.ListIndex = 2 Then
        s_Sql = s_Sql & "FROM plctsresultado res "
        s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      Else
        s_Sql = s_Sql & "FROM plpvs" & IIf(cmbProceso.ListIndex = 0, "vacaciondet", "gratifica") & " res "
      End If
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.pdoano=pdo.anopdo AND res.pdomes=pdo.mespdo AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON pdo.codcls=dxr.codcls AND pdo.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON pdo.codcls=dxc.codcls AND pdo.codpdo=dxc.codpdo AND res.codpsn=dxc.codpsn "
      s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
      s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
      ' filtrado por seccion
      If cboSeccion.ListIndex <> 0 Then
        s_Sql = s_Sql & "AND dxr.codsec='" & Trim(Left(cboSeccion.Text, 2)) & "' "
      End If
      s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
      s_Sql = s_Sql & "AND res.codmon='" & Right(UCase(sMoneda), 1) & "' "
      s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", dxc.codcco, dxr.codsec, dxr.codubica "
      s_Sql = s_Sql & "HAVING (debe_mn <> 0.00 OR haber_mn <> 0.00 OR debe_me <> 0.00 OR haber_me <> 0.00)"
      If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
    End If
  Next n_Index
  ' Quinto Paso: Debe importe negativo Haber
  s_Sql = "UPDATE " & sArchivo & " SET "
  s_Sql = s_Sql & "debe_mn=ABS(haber_mn), debe_me=ABS(haber_me), "
  s_Sql = s_Sql & "haber_mn=0.00, haber_me=0.00 "
  s_Sql = s_Sql & "WHERE haber_mn<0"
  If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub
  ' Sexto Paso: Haber importe negativo Debe
  s_Sql = "UPDATE " & sArchivo & " SET "
  s_Sql = s_Sql & "haber_mn=ABS(debe_mn), haber_me=ABS(debe_me), "
  s_Sql = s_Sql & "debe_mn=0.00, debe_me=0.00 "
  s_Sql = s_Sql & "WHERE debe_mn<0"
  If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then Exit Sub

End Sub

Private Sub cmbPeriodo_Click()
  Dim dFecha As String
  
  cboPeriodo.Clear
  cboPeriodo.AddItem "< Todos >"
  If Trim(cmbPeriodo.Text) <> "" Then
    dFecha = Format(gdl_Funcion.NumeroDiasMes(Left(cmbPeriodo.Text, 2), ps_Anyo), "00")
    dFecha = dFecha & "/" & Left(cmbPeriodo.Text, 2) & "/" & ps_Anyo
    gdl_Procedure.EditDTPicker "AT", dtpFecha, dFecha, s_MdoData_Ins, True, s_FormatoFecha, dtpShortDate
    If s_OptRegistro = "pllasconta" Then
      s_Sql = "SELECT codpdo, despdo "
      s_Sql = s_Sql & "FROM plperiodo "
      s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "'"
      s_Sql = s_Sql & "AND mespdo='" & Left(cmbPeriodo.Text, 2) & "'"
      s_Sql = s_Sql & "AND estadopdo IN('" & s_Estado_Act & "', '" & s_Estado_Blq & "') "
      s_Sql = s_Sql & "ORDER BY codpdo"
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
        While Not porstRecordset.EOF
          cboPeriodo.AddItem gdl_Funcion.PadR(porstRecordset!codpdo, 8, " ") & " - " & porstRecordset!despdo
          porstRecordset.MoveNext
        Wend
        porstRecordset.Close
      End If
    End If
  End If
  gdl_Procedure.EditCombo "PK", cboPeriodo, 0, IIf(s_OptRegistro = "pllasconta", s_MdoData_Ins, s_MdoData_Upd), False
  
End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim sArchivo As String, sFile As String
  Dim sExpresion As String
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Genero tabla temporal
  sArchivo = "tmp" & Format(Now, "yyyymmddhhmmss")
  If Index <> 4 Then
    s_Sql = "CREATE TABLE IF NOT EXISTS " & sArchivo & " ( "
    s_Sql = s_Sql & "codcta varchar(15) Not Null, codpsn varchar(15) Null, codref varchar(11) Null, codcco varchar(10) Null, "
    s_Sql = s_Sql & "codsec char(2) Null, codubica char(2) Null, detalle varchar(60) Null, codmon char(1) Null, "
    s_Sql = s_Sql & "debe_mn decimal(18,2) Null Default '0', haber_mn decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "debe_me decimal(18,2) Null Default '0', haber_me decimal(18,2) Null Default '0') "
    If Not gdl_Funcion.Execution(ps_StrgConnec & ps_DataBase, s_Sql) Then GoTo Finalizar
  End If
  Select Case Index
   Case 0     ' Actualizo detalle de comprobante
    ' Inicializo los totales
    lblTotales(0) = "0.00": lblTotales(1) = "0.00"
    lblTotales(2) = "0.00": lblTotales(3) = "0.00"
    ' Realizo las validaciones de los campos a actualizar
    If cmbPeriodo.Text <> "" Then RecuperaRegistros sArchivo
    ' Recupera información
    s_Sql = "SELECT tmp.codcta, cta.detcta, tmp.codpsn, tmp.codref, tmp.codcco, tmp.codmon, "
    s_Sql = s_Sql & "tmp.debe_mn, tmp.haber_mn, tmp.debe_me, tmp.haber_me "
    s_Sql = s_Sql & "FROM " & sArchivo & " tmp "
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
    s_Sql = s_Sql & "ORDER BY codcta"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    
    tdbRegistro.DataSource = porstRecordset
    ' Obtengo los totales
    s_Sql = "SELECT ROUND(IFNULL(SUM(IFNULL(tmp.debe_mn, 0)), 0), 2) AS debe, ROUND(IFNULL(SUM(IFNULL(tmp.haber_mn, 0)), 0), 2) AS haber, "
    s_Sql = s_Sql & "ROUND(IFNULL(SUM(IFNULL(tmp.debe_me, 0)), 0), 2) AS cargo, ROUND(IFNULL(SUM(IFNULL(tmp.haber_me, 0)), 0), 2) AS abono "
    s_Sql = s_Sql & "FROM " & sArchivo & " tmp "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    lblTotales(0) = FormatNumber(CDec(porstRecordset!debe), 2)
    lblTotales(1) = FormatNumber(CDec(porstRecordset!haber), 2)
    lblTotales(2) = FormatNumber(CDec(porstRecordset!Cargo), 2)
    lblTotales(3) = FormatNumber(CDec(porstRecordset!abono), 2)
   Case 1     ' Genero voucher de contabilidad
    ' Verifico que periodo no se encuentre bloqueado, contabilizar
    If cboParametro.ListIndex <> 0 Then Beep: MsgBox "Proceso bloqueado para contabilizar comprobante", vbExclamation: Exit Sub
    If ps_EmpresaCon = "" Then Beep: MsgBox "Codigo de Empresa de Contabilidad No Valido", vbExclamation: Exit Sub
    ' Realizo las validaciones
    If txtDiario.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Diario de Comprobante", vbExclamation: txtDiario.SetFocus: GoTo Finalizar
    If txtComprobante.Text = "" Then Beep: MsgBox "Debe Ingresar el Numero de Comprobante", vbExclamation: txtComprobante.SetFocus: GoTo Finalizar
    If txtGlosa.Text = "" Then Beep: MsgBox "Debe Ingresar la Glosa del Comprobante", vbExclamation: txtGlosa.SetFocus: GoTo Finalizar
    If cmbPeriodo.Text = "" Then Beep: MsgBox "Debe Selecionar Mes de Proceso", vbInformation: cmbPeriodo.SetFocus: GoTo Finalizar
    If Not (Trim(dtpFecha.Year) = ps_Anyo) Then Beep: MsgBox "Fecha debe ser del ejercico de Transferencia", vbCritical: dtpFecha.SetFocus: GoTo Finalizar
    If Format(dtpFecha.Month, "00") <> Left(Trim(cmbPeriodo.Text), 2) Then Beep: MsgBox "Fecha debe ser del mes del Comprobante", vbCritical: dtpFecha.SetFocus: GoTo Finalizar
    If cmbProceso.Text = "" Then Beep: MsgBox "Debe Seleccionar Proceso Cálculo", vbInformation: cmbProceso.SetFocus: GoTo Finalizar
    ' Verifico periodo bloqueado
    sExpresion = s_Estado_Act
    s_Sql = "SELECT IFNULL(cfg.indcpb, 0) AS indcpb "
    s_Sql = s_Sql & "FROM sysmacon.cociemes cfg "
    s_Sql = s_Sql & "WHERE cfg.codemp='" & ps_EmpresaCon & "' "
    s_Sql = s_Sql & "AND cfg.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND cfg.mescie='" & Left(cmbPeriodo.Text, 2) & "'"
    s_ConexiConta = Mid(ps_StrgConnec, InStr(ps_StrgConnec, "server="), (InStr(InStr(ps_StrgConnec, "server="), ps_StrgConnec, ";") - InStr(ps_StrgConnec, "server=")))
    s_ConexiConta = Replace(ps_StrgConnec, s_ConexiConta, "server=" & ps_ServidorCon)
    Set porstRecordset = OpenRecordset(s_ConexiConta & "sysmacon", adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.EOF) Then
      sExpresion = porstRecordset!indcpb
    End If
    Set porstRecordset = Nothing
    If sExpresion = s_Estado_Act Then Beep: MsgBox "Periodo bloqueado para contabilizar comprobante", vbExclamation: Exit Sub
    ' If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable & " para contabilizar", vbExclamation: Exit Sub
    If MsgBox("¿ Estás Seguro de Contabilizar Comprobante (" & txtDiario.Text & "-" & txtComprobante.Text & ") ?", vbQuestion + vbYesNo) = vbYes Then
      ContabilizaComprobante sArchivo
      MsgBox "Proceso de Exportación Finalizo con Exito", vbInformation
    End If
   Case 2     ' Genera archivo informacion a contabilidad
    ' Realizo las validaciones
    If txtDiario = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Diario de Comprobante", vbExclamation: txtDiario.SetFocus: GoTo Finalizar
    If txtComprobante = "" Then Beep: MsgBox "Debe Ingresar el Numero de Comprobante", vbExclamation: txtComprobante.SetFocus: GoTo Finalizar
    If txtGlosa = "" Then Beep: MsgBox "Debe Ingresar la Glosa del Comprobante", vbExclamation: txtGlosa.SetFocus: GoTo Finalizar
    If cmbPeriodo.Text = "" Then Beep: MsgBox "Debe Selecionar Mes de Proceso", vbInformation: cmbPeriodo.SetFocus: GoTo Finalizar
    If Not (Trim(dtpFecha.Year) = ps_Anyo) Then Beep: MsgBox "Fecha debe ser del ejercico de Transferencia", vbCritical: dtpFecha.SetFocus: GoTo Finalizar
    If Format(dtpFecha.Month, "00") <> Left(Trim(cmbPeriodo.Text), 2) Then Beep: MsgBox "Fecha debe ser del mes del Comprobante", vbCritical: dtpFecha.SetFocus: GoTo Finalizar
    If cmbProceso.Text = "" Then Beep: MsgBox "Debe Seleccionar Proceso Cálculo", vbInformation: cmbProceso.SetFocus: GoTo Finalizar
    
    sExpresion = IIf((cboParametro.ListIndex = 3 Or cboParametro.ListIndex = 6 Or cboParametro.ListIndex = 7 Or cboParametro.ListIndex = 8), ".xlsx", ".txt")
    sFile = Trim(ps_RucEmpresa) & "rd" & ps_Anyo & Left(cmbPeriodo.Text, 2) & sExpresion
    On Error GoTo CancelaDialogo
    fMenu.cdlDialogo.DialogTitle = "Grabar Archivo Como"
    fMenu.cdlDialogo.CancelError = True
    fMenu.cdlDialogo.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    fMenu.cdlDialogo.FileName = sFile
    fMenu.cdlDialogo.DefaultExt = sExpresion
    sExpresion = "Archivo de Excel (*.xls)|*.xlsx|Archivo de Excel 97-2003 (*.xls)|*.xls|Todos los archivos(*.*)|*.*"
    sExpresion = IIf((cboParametro.ListIndex = 3 Or cboParametro.ListIndex = 6 Or cboParametro.ListIndex = 7), sExpresion, "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*")
    fMenu.cdlDialogo.Filter = sExpresion
    fMenu.cdlDialogo.ShowSave
  
CancelaDialogo:
    ' verifico si existe error y desactivo
    If Not Err.Number = 0 Then MsgBox Error(Err.Number): GoTo Finalizar
    On Error GoTo 0
    
    ChDir App.path
    If MsgBox("¿ Estás Seguro de Generar Archivo de Comprobante Contable? ", vbQuestion + vbYesNo) = vbYes Then
      sFile = fMenu.cdlDialogo.FileName
      Select Case cboParametro.ListIndex
       Case 1
        GenArchivoSap IIf(chkProceso.Value, "_pdo", ""), sFile
       Case 2
        GenArchivoSpring sArchivo, sFile, "G"
       Case 3
        GenArchivoSapExcel sArchivo, sFile, "G"
       Case 4
        GenArchivoOracle sArchivo, sFile, "G"
       Case 5
        GenArchivoSpring_Proyecto sArchivo, sFile, "G"
       Case 6
        GenArchivoConcarExcel sArchivo, sFile, "G"
       Case 7
        GenArchivoInformatExcel sArchivo, sFile, "G"
       Case 8
         GenArchivoPLCentroExcel sArchivo, sFile, "G"
       Case Else
        ExportaContabilidad sArchivo, sFile, "G"
      End Select
      If s_StatusValid_DatosConcar = "OK" Then
        MsgBox "Proceso de Exportación Finalizo con Exito", vbInformation
      End If
    End If
    ChDrive Left$(App.path, 1)
    ChDir App.path
    
   Case 3, 4  ' Opciones de impresión
    ' Realizo las validaciones
    If txtDiario = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Diario de Comprobante", vbExclamation: txtDiario.SetFocus: GoTo Finalizar
    If txtComprobante = "" Then Beep: MsgBox "Debe Ingresar el Numero de Comprobante", vbExclamation: txtComprobante.SetFocus: GoTo Finalizar
    If txtGlosa = "" Then Beep: MsgBox "Debe Ingresar la Glosa del Comprobante", vbExclamation: txtGlosa.SetFocus: GoTo Finalizar
    If cmbPeriodo.Text = "" Then Beep: MsgBox "Debe Selecionar Mes de Proceso", vbInformation: cmbPeriodo.SetFocus: GoTo Finalizar
    If Not (Trim(dtpFecha.Year) = ps_Anyo) Then Beep: MsgBox "Fecha debe ser del ejercico de Transferencia", vbCritical: dtpFecha.SetFocus: GoTo Finalizar
    If Format(dtpFecha.Month, "00") <> Left(Trim(cmbPeriodo.Text), 2) Then Beep: MsgBox "Fecha debe ser del mes del Comprobante", vbCritical: dtpFecha.SetFocus: GoTo Finalizar
    If cmbProceso.Text = "" Then Beep: MsgBox "Debe Seleccionar Proceso Cálculo", vbInformation: cmbProceso.SetFocus: GoTo Finalizar
        
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = s_TitleWindow
    gdl_Procedure.ps_ReportName = "rpttransconta"
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & "COMPROBANTE CONTABLE" & " - " & ps_DesClsPlanilla & ";true"
    aElemento(2, 2) = "Periodo;" & Mid(cmbPeriodo.Text, 6) & " - " & ps_Anyo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "diario varchar(4) NOT Null, comprobante varchar(6) NOT Null, "
    s_Sql = s_Sql & "fecha date NOT Null, glosa varchar(60) Null, "
    s_Sql = s_Sql & "codcta varchar(15) NOT Null, codpsn varchar(15) Null, "
    s_Sql = s_Sql & "codcco varchar(10) Null, detalle varchar(60) Null, "
    s_Sql = s_Sql & "codmon char(1) Null, tipcambio decimal(7,3) Null Default '0', "
    s_Sql = s_Sql & "debe_mn decimal(18,2) Null Default '0', haber_mn decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "debe_me decimal(18,2) Null Default '0', haber_me decimal(18,2) Null Default '0') "
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Genera la información del reporte
    ExportaContabilidad sArchivo, "tmp" & gdl_Procedure.ps_ReportName, "R"
    ' Recupera información
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY codcta"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 3), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' ]
  End Select

Finalizar:
  ' Elimino la tabla temporal
  s_Sql = "DROP TABLE IF EXISTS " & sArchivo
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal

End Sub

Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim Item As New ValueItem   ' Cambio el formato de la grilla columna de valores
  'ABRIL 2015
  'Estatus valida Concar
  s_StatusValid_DatosConcar = "OK"
  
  ' Establece posición del formulario
  Me.Height = 7530: Me.Width = 9200
  Me.Left = 1250: Me.Top = 150
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  s_OptRegistro = s_SwRegistro
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Contabilización de " & IIf(s_OptRegistro = "pllasconta", "Remuneraciones", "Provisiones")
  s_TitleTable = "Detalle Comprobante"
  
  ReDim aElemento(9, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Cuenta", "Descripción", "Personal", "Cen.Costo", "Mon", "Debe MN", "Haber MN", "Debe ME", "Haber ME")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcta", "detcta", "codpsn", "codcco", "codmon", "debe_mn", "haber_mn", "debe_me", "haber_me")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1000, 2006.03, 810, 900, 450, 950, 950, 950, 950)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "", "standard", "standard", "standard", "standard")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "", "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0, 1, 1, 1, 1)
  Next n_Index
  ReDim aElementos(2, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 3
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(4).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(4).ValueItems.Validate = True
  tdbRegistro.Columns(4).ValueItems.Translate = True
  tdbRegistro.Columns(4).ValueItems.CycleOnClick = True
  For n_Index = 0 To 1
    tdbRegistro.Columns(4).ValueItems.Add Item
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Codmon_mn, s_Codmon_me)
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, s_Codmon_mn_Txt, s_Codmon_me_Txt)
  Next n_Index
  ']
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(5, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "proceso": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "detactas", "asoclibr", "genarchi", "prelimin", "imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Actualiza Detalle Comprobante", "Genera Voucher de Contabilidad", "Genera Archivo de Transferencia", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo el grafico del boton de seccion
  
  For n_Index = 0 To 8
    cboParametro.AddItem "Sistema " & Choose(n_Index + 1, "General", "SAP", "SPRING", "SAP - Excel", "ORACLE", "SPRING - Proyecto", "CONCAR - Excel", "INFORMAT - Excel", "PLCENTRO - Excel")
  Next n_Index
  gdl_Procedure.EditCombo "PK", cboParametro, 0, s_MdoData_Ins, False
  
  ' Carga los datos en el formulario
  gdl_Procedure.EditText "AT", txtDiario, "", s_MdoData_Ins, False, 4, vbLeftJustify
  gdl_Procedure.EditText "AT", txtComprobante, Format("0", "000000"), s_MdoData_Ins, False, 6, vbLeftJustify
  gdl_Procedure.EditText "AT", txtGlosa, "", s_MdoData_Ins, False, 60, vbLeftJustify
  gdl_Procedure.EditOptionCheck "AT", chkProceso, False, s_MdoData_Ins, True
  
  cmbPeriodo.Clear
  cmbPeriodo.Locked = False
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  ' Periodo de pago
  cboPeriodo.Clear
  cboPeriodo.AddItem "< Todos >"
  gdl_Procedure.EditCombo "PK", cboPeriodo, 0, IIf(s_OptRegistro = "pllasconta", s_MdoData_Ins, s_MdoData_Upd), False
  
  ' seccion empresa
  cboSeccion.Clear
  cboSeccion.AddItem "< Todos >"
  s_Sql = "SELECT codsec, dessec "
  s_Sql = s_Sql & "FROM plseccion "
  s_Sql = s_Sql & "WHERE estadosec IN('" & s_Estado_Act & "', '" & s_Estado_Blq & "') "
  s_Sql = s_Sql & "ORDER BY codsec"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    While Not porstRecordset.EOF
      cboSeccion.AddItem gdl_Funcion.PadR(porstRecordset!codsec, 2, " ") & " - " & porstRecordset!dessec
      porstRecordset.MoveNext
    Wend
    porstRecordset.Close
  End If
  gdl_Procedure.EditCombo "PK", cboSeccion, 0, s_MdoData_Ins, False
  
  ' proceso de Cálculo
  cmbProceso.Clear
  cmbProceso.Locked = False
  If s_OptRegistro = "pllasconta" Then
    s_Sql = "SELECT codproce, desproce "
    s_Sql = s_Sql & "FROM plproceso "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND estadoproce<>'" & s_EstadoRemAper & "' "
    s_Sql = s_Sql & "ORDER BY codproce"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      Do While Not porstRecordset.EOF
        cmbProceso.AddItem porstRecordset("desproce") & Space(60) & "|" & Trim(porstRecordset("codproce"))
        porstRecordset.MoveNext
      Loop
      porstRecordset.Close
      cmbProceso.ListIndex = 0
    End If
    cmbProceso.Width = 4110
  ElseIf s_OptRegistro = "pvscontabi" Then
    For n_Index = 1 To 3
      cmbProceso.AddItem "Provisión de " & Choose(n_Index, "Vacaciones", "Gratificaciones", "C.T.S.")
    Next n_Index
    cmbProceso.ListIndex = 0
    cmbProceso.Width = 3000
  End If
  gdl_Procedure.EditDTPicker "AT", dtpFecha, Date, s_MdoData_Ins, True, s_FormatoFecha, dtpShortDate
  ']
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0

  ' Actualiza información de comprobante
  cmdAction_Click s_Estado_Ina
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
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

Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then cmdAction_Click s_Estado_Ina
End Sub
Private Sub txtComprobante_GotFocus()
  gdl_Procedure.MarcaGet txtComprobante
End Sub
Private Sub txtDiario_GotFocus()
  gdl_Procedure.MarcaGet txtDiario
End Sub
Private Sub txtGlosa_GotFocus()
  gdl_Procedure.MarcaGet txtGlosa
End Sub
