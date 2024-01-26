VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fReporPlanillAfp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "rptplanillafp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   6135
   Begin VB.DirListBox Dir 
      Height          =   2340
      Left            =   1560
      TabIndex        =   25
      Top             =   585
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exportación de Datos AFPNET"
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
      Height          =   1080
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cmbmes 
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
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "rptplanillafp.frx":000C
         Left            =   1560
         List            =   "rptplanillafp.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   2625
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin Threed.SSCommand cmdexportar 
         Height          =   360
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "rptplanillafp.frx":0010
      End
      Begin Threed.SSCommand cmdbusqueda 
         Height          =   360
         Left            =   5715
         TabIndex        =   23
         Top             =   240
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   0
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "rptplanillafp.frx":0692
      End
      Begin Threed.SSCommand cmdexcel 
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         BevelWidth      =   0
         AutoSize        =   2
         Picture         =   "rptplanillafp.frx":0D14
      End
      Begin Threed.SSCommand cmdverificar 
         Height          =   300
         Left            =   5760
         TabIndex        =   32
         Top             =   600
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         BevelWidth      =   0
         AutoSize        =   2
         Picture         =   "rptplanillafp.frx":12AE
      End
      Begin VB.Label lblDato 
         Caption         =   "Mes :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   900
      End
      Begin VB.Label ruta 
         BackColor       =   &H80000016&
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1560
         TabIndex        =   26
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label labelruta 
         BackColor       =   &H80000013&
         Caption         =   "Grabar en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4605
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   8123
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
      Left            =   0
      Top             =   6720
      Width           =   5250
      _ExtentX        =   9260
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
      Height          =   5025
      Index           =   0
      Left            =   5355
      TabIndex        =   5
      Top             =   2040
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8864
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
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Registro"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   96
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
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1500
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
         Picture         =   "rptplanillafp.frx":1408
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1920
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
         Picture         =   "rptplanillafp.frx":1424
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2580
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
         Picture         =   "rptplanillafp.frx":1440
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3015
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
         Picture         =   "rptplanillafp.frx":145C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   13
         Tag             =   "0"
         Top             =   4035
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
         Picture         =   "rptplanillafp.frx":1478
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   14
         Tag             =   "0"
         Top             =   4470
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
         Picture         =   "rptplanillafp.frx":1494
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1065
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
         Picture         =   "rptplanillafp.frx":14B0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   3435
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
         Picture         =   "rptplanillafp.frx":14CC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   6
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
         Picture         =   "rptplanillafp.frx":14E8
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   1060
      Width           =   6045
      _Version        =   65536
      _ExtentX        =   10663
      _ExtentY        =   1693
      _StockProps     =   14
      Caption         =   " Parametro de Selección del Reporte"
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
      Begin VB.CheckBox chkfirma 
         Caption         =   "Firma "
         Height          =   195
         Left            =   5160
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox s27252 
         Caption         =   "Sin  Ley 27252"
         Height          =   255
         Left            =   4560
         TabIndex        =   31
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox c27252 
         Caption         =   "Con Ley 27252"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   240
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "rptplanillafp.frx":1504
         Left            =   1155
         List            =   "rptplanillafp.frx":1506
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   2625
      End
      Begin VB.TextBox txtAfp 
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
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   4
         Top             =   600
         Width           =   705
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   17
         Top             =   600
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
         Left            =   2325
         TabIndex        =   18
         Top             =   645
         Width           =   195
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Entidad :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Mes :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   900
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1560
      TabIndex        =   19
      Top             =   1800
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
Attribute VB_Name = "fReporPlanillAfp"
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
Private cnn As ADODB.Connection
Private valor As String
Private mesanterior As String

Function Recordset_a_Csv(rs As Recordset, path As String) As Boolean
  Dim cadenav As String, ceros As String
  Dim Fila As Integer, desde As Integer
  Dim tamanno As Integer, posicion As Integer
  Dim Columna As Long
  
  On Error GoTo Err_function
  ' Crea el archivo
  Open path For Output As #1
  ' Se mueve al primer registro
  rs.MoveFirst
  
  ' recorre todo el recordset
  For Fila = 0 To rs.RecordCount - 1
    ' nombre del campo
    cadenav = ""
    tamanno = 12
    valor = Trim(rs.Fields(0))
    
    If Len(valor) = tamanno Then
    ElseIf Len(valor) > tamanno Then
      valor = Left(valor, tamanno)
    ElseIf Len(valor) < tamanno Then
      cadenav = gdl_Funcion.PadR(cadenav, (tamanno - Len(valor)), " ")
    End If
    valor = valor & cadenav
  
    Print #1, Format(Fila + 1, "00000") & valor;
    ' recorre todos los campos
    For Columna = 1 To rs.Fields.Count - 1
      ' imprime la fila actual en el fichero
      valor = IIf(IsNull(rs.Fields(Columna)), "", Trim(rs.Fields(Columna)))
      cadenav = ""
      
      Select Case Columna
       Case 1, 6, 7, 8, 9, 14
        tamanno = 1
       Case 2, 3, 4, 5
        tamanno = 20
       Case 10, 11, 12, 13
        tamanno = 9
       Case 15
        tamanno = 2
      End Select
    
      If Len(valor) = tamanno Then
      ElseIf Len(valor) > tamanno Then
        valor = Left(valor, tamanno)
      ElseIf Len(valor) < tamanno Then
        If Columna <> 10 Then
          cadenav = gdl_Funcion.PadR(cadenav, (tamanno - Len(valor)), " ")
        End If
      End If
    
      If Columna = 10 Then
        posicion = InStr(valor, ".")
        If Len(valor) - posicion = 1 Then
          valor = valor & "0"
        End If
        If posicion = 0 Then
          valor = valor & ".00"
        End If
        posicion = InStr(valor, ".")
        While 6 > posicion - 1
          ceros = "0" & ceros
          posicion = posicion + 1
        Wend
        valor = ceros & valor
        ceros = ""
      End If
    
      If (Columna >= 11 And Columna <= 13) Then
        cadenav = "000000.00"
      End If
      valor = valor & cadenav
      Print #1, "" & valor;
    Next
    ' escribe una línea en blanco
    Print ""
    ' salto de carro
    Print #1, "" & Chr(13) & Chr(10);
    ' mueve el recordset al siguiente registro
    rs.MoveNext
  Next
  ' cierra el archivo
  Close #1
  Exit Function

Err_function:
  MsgBox Err.Description, vbCritical
  Close
End Function

'[
Private Sub PlanillaPrevision(ByVal s_Archivo As String, ByVal s_Periodo As String, s_Proceso As String, s_FechaHora As String, s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, a_Parametro(8) As String
  Dim s_Caracter As String, s_Trabajador As String
  Dim s_Parametro(8) As String, n_Importe As Double
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sDireccion As String, sNumero As String, sLocalidad As String
  Dim sDepartamento As String, sProvincia As String, sDistrito As String
  Dim sCodPostal As String, sTelefono As String
  Dim sRepresentante As String, sTipoDocu As String, sDocumento As String
  Dim sEncargado As String, sEmail As String, sCencosto As String, sAnexo As String
  Dim sDireccionpsn As String, nPersonal As Long, nDesafiliado As Long
  
  ' Recupero los parametros de configuración
  s_Sql = "SELECT prm.cpcremase, prm.cpcapobli, prm.cpcapovolsfp, prm.cpcapovolcfp, prm.cpcapoemp, prm.cpcseguro, prm.cpcporcen,prm.cpc27252, "
  s_Sql = s_Sql & "cfg.codvia, via.abrevia, cfg.direccionvia, cfg.numerodir, cfg.codzona, zon.abrezona, cfg.direccionzona, cfg.ubigeodir, cfg.telefono, "
  s_Sql = s_Sql & "cfg.email, CONCAT(IFNULL(cfg.repapepaterno, ''), ' ', IFNULL(cfg.repapematerno, ''), ', ', IFNULL(cfg.repnombres, '')) AS representante, cfg.repcoddci, cfg.repnumdocu, "
  s_Sql = s_Sql & "CONCAT(IFNULL(cfg.psnapepaterno, ''), ' ', IFNULL(cfg.psnapematerno, ''), ', ', IFNULL(cfg.psnnombres, '')) AS encargado, cfg.psntelefono, cfg.codcco "
  s_Sql = s_Sql & "FROM plparametroafp prm "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON prm.pdoano=cfg.pdoano "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON cfg.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON cfg.codzona=zon.codzona "
  s_Sql = s_Sql & "WHERE prm.pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de la planilla", vbCritical: Exit Sub
  If gdl_Funcion.aTexto(porstRecordset!cpcremase) = "" Then Beep: MsgBox "Debe configurar los parametros de la planilla", vbCritical: Exit Sub
  
  ' Obtengo los conceptos de recuperación
  a_Parametro(1) = gdl_Funcion.aTexto(porstRecordset!cpcremase)
  a_Parametro(2) = gdl_Funcion.aTexto(porstRecordset!cpcapobli)
  a_Parametro(3) = gdl_Funcion.aTexto(porstRecordset!cpcapovolsfp)
  a_Parametro(4) = gdl_Funcion.aTexto(porstRecordset!cpcapovolcfp)
  a_Parametro(5) = gdl_Funcion.aTexto(porstRecordset!cpcapoemp)
  a_Parametro(6) = gdl_Funcion.aTexto(porstRecordset!cpcseguro)
  a_Parametro(7) = gdl_Funcion.aTexto(porstRecordset!cpcporcen)
  a_Parametro(8) = gdl_Funcion.aTexto(porstRecordset!cpc27252)
  ' Datos de la cabecera
  sDireccion = gdl_Funcion.aTexto(porstRecordset!abrevia)
  sDireccion = sDireccion & " " & gdl_Funcion.aTexto(porstRecordset!direccionvia)
  sNumero = "Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir)
  sLocalidad = gdl_Funcion.aTexto(porstRecordset!abrezona)
  sLocalidad = sLocalidad & " " & gdl_Funcion.aTexto(porstRecordset!direccionzona)
  sDistrito = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
  sDepartamento = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, "0", Left(sDistrito, 2), "UB")
  sProvincia = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, "1", Left(sDistrito, 4), "UB")
  sDistrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDistrito, "UB")
  sCodPostal = Left(gdl_Funcion.aTexto(porstRecordset!telefono), 2)
  sTelefono = Mid(gdl_Funcion.aTexto(porstRecordset!telefono), 4)
  sRepresentante = gdl_Funcion.aTexto(porstRecordset!representante)
  sTipoDocu = gdl_Funcion.aTexto(porstRecordset!repcoddci)
  sDocumento = gdl_Funcion.aTexto(porstRecordset!repnumdocu)
  sEncargado = gdl_Funcion.aTexto(porstRecordset!encargado)
  sEmail = gdl_Funcion.aTexto(porstRecordset("email"))
  sCencosto = gdl_Funcion.aTexto(porstRecordset!codcco)
  sCencosto = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, ps_CodEmpresa, sCencosto, "CC")
  sAnexo = gdl_Funcion.aTexto(porstRecordset!psntelefono)
  porstRecordset.Close
  
  ' Recupero la información para exportar
  s_Sql = "SELECT DISTINCTROW res.codcls, res.codpsn, res.codcpc, psn.numeroafp, psn.apepaterno, psn.apematerno, psn.nombres, "
  s_Sql = s_Sql & "dxr.fecingreso, MAX(dxr.estadopsn) AS estadopsn, MAX(dxr.fecestado) AS fecestado, "
  s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, zon.abrezona, psn.nomzondirec, "
  s_Sql = s_Sql & "psn.ubigeodir, psn.telefono, "
  s_Sql = s_Sql & "pll.nrohoja, pll.sinpago, pll.fechapago, pll.formapago, "
  s_Sql = s_Sql & "pll.interespension, pll.chequepension, pll.codbcopension, "
  s_Sql = s_Sql & "pll.interesadmin, pll.chequeadmin, pll.codbcoadmin, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0)), 2) AS importe "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn AND dxr.codafp='" & Trim(txtAfp.Text) & "' "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plplanillafp pll ON res.pdoano=pll.pdoano AND res.pdomes=pll.pdomes AND pll.codafp='" & Trim(txtAfp.Text) & "' "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
  s_Sql = s_Sql & "WHERE res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.pdomes='" & s_Periodo & "' "
  
  If c27252.Value = Checked Then
    s_Sql = s_Sql & "AND psn.chk27252=1 "
  End If
  
  If s27252.Value = Checked Then
    s_Sql = s_Sql & "AND psn.chk27252=0 "
  End If
  
  s_Sql = s_Sql & "AND res.codcls IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.codcpc IN ('" & a_Parametro(1) & "', '" & a_Parametro(2) & "', '" & a_Parametro(3) & "', '" & a_Parametro(4) & "', '" & a_Parametro(5) & "', '" & a_Parametro(6) & "', '" & a_Parametro(7) & "', '" & a_Parametro(8) & "') "
  s_Sql = s_Sql & "GROUP BY res.codcls, res.codpsn, res.codcpc "
  s_Sql = s_Sql & "ORDER BY res.codcls, res.codpsn, res.codcpc"
  
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
     ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Proceso Planilla AFP ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    nPersonal = 0: nDesafiliado = 0
    n_Index = 0
    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("codpsn", "numeroafp", "apepaterno", "apematerno", "nombres", "fecingregpen", "direccionpsn", "telefonopsn", "estadopsn", "fecestado", "direccion", "numero", "localidad", "departamento", "provincia", "distrito", "postal", "telefono", "representante", "tipodocu", "documento", "encargado", "email", "cencosto", "anexo", "moneda", "nrohoja", "nroafiliados", "fechapago", "formapago", "interespension", "chequepension", "bcopension", "interesadmin", "chequeadmin", "bcoadmin", "impremase", "impapobli", "impapovolsfp", "impapovolcfp", "impapoemp", "impseguro", "impporcen", "ejercicio", "periodo", "afiliado", "codcls", "imp27252")
      a_Valores = Array("", "", "", "", "", "", "", "", "", "", sDireccion, sNumero, sLocalidad, sDepartamento, sProvincia, sDistrito, sCodPostal, sTelefono, sRepresentante, sTipoDocu, sDocumento, sEncargado, sEmail, sCencosto, sAnexo, IIf(fMenu.ribMoneda(0).Value, "N", "E"), "", CLng(0), "", "", CDec(0), "", "", CDec(0), "", "", CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), ps_Anyo, s_Periodo, 0, "", CDec(0))
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
      s_Caracter = "|"
    End If
    While Not porstRecordset.EOF
      ' Personal afiliados y desafiliados
      nPersonal = nPersonal + 1
      n_Index = 0
      If (Format(porstRecordset("fecestado"), "yyyymm") = ps_Anyo & s_Periodo And Trim(porstRecordset("estadopsn")) = "I") Then
        nDesafiliado = nDesafiliado + 1
        n_Index = 1
      End If
      ' Genero el registro de grabación
      psRegistro = ""
      psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codpsn) & s_Caracter
      
      s_Trabajador = porstRecordset!codpsn
      s_Parametro(1) = "": s_Parametro(2) = "": s_Parametro(3) = ""
      s_Parametro(4) = "": s_Parametro(5) = "": s_Parametro(6) = "": s_Parametro(7) = "": s_Parametro(8) = ""
      ' Obtengo la direccion del personal
      sDireccionpsn = gdl_Funcion.aTexto(porstRecordset("abrevia"))
      sDireccionpsn = sDireccionpsn & " " & gdl_Funcion.aTexto(porstRecordset("nomviadirec"))
      If Not IsNull(porstRecordset("numerdirec")) Then
        sDireccionpsn = sDireccionpsn & " Nº " & gdl_Funcion.aTexto(porstRecordset("numerdirec"))
      End If
      If Not IsNull(porstRecordset("intedirec")) Then
        sDireccionpsn = sDireccionpsn & " / " & gdl_Funcion.aTexto(porstRecordset("intedirec"))
      End If
      sDireccionpsn = sDireccionpsn & " " & gdl_Funcion.aTexto(porstRecordset("abrezona"))
      sDireccionpsn = sDireccionpsn & " " & gdl_Funcion.aTexto(porstRecordset("nomzondirec"))
      sDireccionpsn = sDireccionpsn & " " & gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, gdl_Funcion.aTexto(porstRecordset("ubigeodir")), "UB")
      If s_Accion = "R" Then
        a_Valores(0) = porstRecordset("codpsn"): a_Valores(1) = gdl_Funcion.aTexto(porstRecordset("numeroafp"))
        a_Valores(2) = gdl_Funcion.aTexto(porstRecordset("apepaterno")): a_Valores(3) = gdl_Funcion.aTexto(porstRecordset("apematerno"))
        a_Valores(4) = gdl_Funcion.aTexto(porstRecordset("nombres")): a_Valores(5) = Format(porstRecordset("fecingreso"), s_FmtFechMysql_0)
        a_Valores(6) = sDireccionpsn: a_Valores(7) = gdl_Funcion.aTexto(porstRecordset("telefono"))
        a_Valores(8) = porstRecordset("estadopsn"): a_Valores(9) = Format(porstRecordset("fecestado"), s_FmtFechMysql_0)
        a_Valores(26) = Trim(porstRecordset("nrohoja")): a_Valores(27) = CLng(nPersonal)
        a_Valores(28) = Format(porstRecordset("fechapago"), s_FmtFechMysql_0): a_Valores(29) = gdl_Funcion.aTexto(porstRecordset("formapago"))
        a_Valores(30) = CDec(porstRecordset("interespension")): a_Valores(31) = gdl_Funcion.aTexto(porstRecordset("chequepension"))
        a_Valores(32) = UCase(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, gdl_Funcion.aTexto(porstRecordset("codbcopension")), "EB"))
        a_Valores(33) = CDec(porstRecordset("interesadmin")): a_Valores(34) = gdl_Funcion.aTexto(porstRecordset("chequeadmin"))
        a_Valores(35) = UCase(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, gdl_Funcion.aTexto(porstRecordset("codbcoadmin")), "EB"))
        a_Valores(45) = n_Index: a_Valores(46) = porstRecordset("codcls")
      End If
      Do
        n_Importe = CDec(porstRecordset("importe"))
        If a_Parametro(1) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(1) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(2) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(2) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(3) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(3) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(4) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(4) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(5) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(5) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(6) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(6) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(7) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(7) = Format(n_Importe, "###########0.00")
        End If
        If a_Parametro(8) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(8) = Format(n_Importe, "###########0.00")
        End If
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        DoEvents
        porstRecordset.MoveNext
        ' Fin de archivo
        If porstRecordset.EOF Then Exit Do
      Loop While s_Trabajador = porstRecordset!codpsn
      
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores(36) = CDec(IIf(s_Parametro(1) = "", 0, s_Parametro(1)))
        a_Valores(37) = CDec(IIf(s_Parametro(2) = "", 0, s_Parametro(2)))
        a_Valores(38) = CDec(IIf(s_Parametro(3) = "", 0, s_Parametro(3)))
        a_Valores(39) = CDec(IIf(s_Parametro(4) = "", 0, s_Parametro(4)))
        a_Valores(40) = CDec(IIf(s_Parametro(5) = "", 0, s_Parametro(5)))
        a_Valores(41) = CDec(IIf(s_Parametro(6) = "", 0, s_Parametro(6)))
        a_Valores(42) = CDec(IIf(s_Parametro(7) = "", 0, s_Parametro(7)))
        a_Valores(47) = CDec(IIf(s_Parametro(8) = "", 0, s_Parametro(8)))
        
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = psRegistro & s_Parametro(1) & s_Caracter
        psRegistro = psRegistro & s_Parametro(2) & s_Caracter
        psRegistro = psRegistro & s_Parametro(3) & s_Caracter
        psRegistro = psRegistro & s_Parametro(4) & s_Caracter
        psRegistro = psRegistro & s_Parametro(5) & s_Caracter
        psRegistro = psRegistro & s_Parametro(6) & s_Caracter
        psRegistro = psRegistro & s_Parametro(7) & s_Caracter
        psRegistro = psRegistro & s_Parametro(8) & s_Caracter
        potxtFileExp.WriteLine psRegistro
      End If
    Wend
    If s_Accion = "G" Then
      ' Cierro objeto y saco de memoria
      potxtFileExp.Close
    End If
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  Else
    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("codpsn", "numeroafp", "apepaterno", "apematerno", "nombres", "fecingregpen", "direccionpsn", "telefonopsn", "estadopsn", "fecestado", "direccion", "numero", "localidad", "departamento", "provincia", "distrito", "postal", "telefono", "representante", "tipodocu", "documento", "encargado", "email", "cencosto", "anexo", "moneda", "nrohoja", "nroafiliados", "fechapago", "formapago", "interespension", "chequepension", "bcopension", "interesadmin", "chequeadmin", "bcoadmin", "impremase", "impapobli", "impapovolsfp", "impapovolcfp", "impapoemp", "impseguro", "impporcen", "ejercicio", "periodo", "afiliado", "codcls", "imp27252")
      a_Valores = Array("", "", "", "", "", "", "", "", "", "", sDireccion, sNumero, sLocalidad, sDepartamento, sProvincia, sDistrito, sCodPostal, sTelefono, sRepresentante, sTipoDocu, sDocumento, sEncargado, sEmail, sCencosto, sAnexo, IIf(fMenu.ribMoneda(0).Value, "N", "E"), "", CLng(0), "", "", CDec(0), "", "", CDec(0), "", "", CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), ps_Anyo, s_Periodo, 0, "", CDec(0))
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero)
    End If
  End If
  If (nPersonal <= 15 And nDesafiliado > 0 And s_Accion = "R") Then
    a_Valores(1) = "": a_Valores(2) = "": a_Valores(3) = "": a_Valores(4) = ""
    a_Valores(5) = "": a_Valores(6) = "": a_Valores(7) = "": a_Valores(8) = "A"
    a_Valores(9) = "": a_Valores(27) = CLng(nPersonal): a_Valores(36) = CDec(0): a_Valores(37) = CDec(0)
    a_Valores(38) = CDec(0): a_Valores(39) = CDec(0): a_Valores(40) = CDec(0): a_Valores(41) = CDec(0)
    a_Valores(42) = CDec(0): a_Valores(47) = CDec(0):
    nPersonal = nPersonal + 1
    For n_Index = nPersonal To 16
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      a_Valores(0) = "zzzzzzzzz" & Format(n_Index, "00")
      a_Valores(46) = "zz"
      ' Realizo la actualización de los registros
      If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    Next n_Index
  End If
  ' Realizo la actualización de los registros (numero de personal)
  If s_Accion = "R" Then
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    s_Sql = "UPDATE " & s_Archivo & " "
    s_Sql = s_Sql & "SET nroafiliados=" & a_Valores(27)
    If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, descls, clave, estadocls"
  s_Sql = s_Sql & " FROM plclasplan"
  s_Sql = s_Sql & " ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly
  
  ' Inicializo los rangos de impresion
  as_SelRegistro(0) = "": as_SelRegistro(1) = ""
  If dcaRegistro.Recordset.RecordCount > 0 Then
    dcaRegistro.Recordset.MoveLast: as_SelRegistro(1) = dcaRegistro.Recordset.Bookmark
    dcaRegistro.Recordset.MoveFirst: as_SelRegistro(0) = dcaRegistro.Recordset.Bookmark
  End If

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_Proceso As String
  Dim sDireccion As String, sRepresentante As String
  Dim sSinPago As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Actualización de parametros
    If cmbPeriodo = "" Then Beep: MsgBox "Debe seleccionar el mes de la planilla", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    If txtAfp = "" Then Beep: MsgBox "Debe Ingresar el Codigo de la Entidad de pensiones", vbExclamation: txtAfp.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Entidad de pensiones no existe; verifique", vbExclamation: txtAfp.SetFocus: Exit Sub
    Me.Tag = s_MdoData_Vis
    fPrmPlanillaAfp.Show vbModal
   Case 1, 2  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 3 ' Busqueda de registro
    Set go_tdbBusqueda = tdbRegistro
    Set go_dcaBusqueda = dcaRegistro
    gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
    fBusqueda.Show vbModal
   Case 4, 5, 6 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), (Index - 4), s_TitleTable
   Case 7, 8  ' Opciones de impresión
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    ' Verifico que existan parametros de planilla
    s_Sql = "SELECT pll.nrohoja, pll.sinpago, prm.cpcremase, cfg.repnombres "
    s_Sql = s_Sql & "FROM plplanillafp pll "
    s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON pll.pdoano=cfg.pdoano "
    s_Sql = s_Sql & "INNER JOIN plparametroafp prm ON pll.pdoano=prm.pdoano "
    s_Sql = s_Sql & "WHERE pll.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND pll.pdomes='" & Left(cmbPeriodo, 2) & "' "
    s_Sql = s_Sql & "AND pll.codafp='" & Trim(txtAfp.Text) & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de la planilla", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!nrohoja) = "" Or gdl_Funcion.aTexto(porstRecordset!cpcremase) = "" Or gdl_Funcion.aTexto(porstRecordset!repnombres) = "" Then Beep: MsgBox "Debe configurar los parametros de la planilla", vbCritical: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Proceso = "planillafp"
    sSinPago = gdl_Funcion.aTexto(porstRecordset!sinpago)
    sSinPago = IIf(sSinPago = s_Estado_Ina, "N", "S")
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Obtengo los datos d ela empresa
    sDireccion = "": sRepresentante = ""
    s_Sql = "SELECT codvia, direccionvia, numerodir, codzona, direccionzona, ubigeodir, CONCAT(repapepaterno, ' ', repapematerno, ', ', repnombres) AS representante "
    s_Sql = s_Sql & "FROM plcfgempresa "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sRepresentante = gdl_Funcion.aTexto(porstRecordset!representante)
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDireccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " - " & sDireccion
    End If
    porstRecordset.Close
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "PLANILLA PREVISIONAL DE AFP"
    
    gdl_Procedure.ps_ReportName = "rptplanillafp"
    
    If c27252.Value = Checked Then
    gdl_Procedure.ps_ReportName = "rptplanillafp27252"
    End If
  
    If s27252.Value = Checked Then
    gdl_Procedure.ps_ReportName = "rptplanillafp27252"
    End If
    
    ReDim aElemento(3, 5): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = "": aElemento(0, 3) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = ""
    aElemento(1, 2) = "": aElemento(1, 3) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    
    aElemento(2, 1) = "TituloReporte;" & "DE APORTES PREVISIONALES" & ";true"
    
    If c27252.Value = Checked Then
        aElemento(2, 1) = "TituloReporte;" & "DE APORTES PREVISIONALES - LEY Nº 27252" & ";true"
    End If
    
    If s27252.Value = Checked Then
        aElemento(2, 1) = "TituloReporte;" & "DE APORTES PREVISIONALES - LEY Nº 27252" & ";true"
    End If
     
    aElemento(2, 2) = "Ruc;" & ps_RucEmpresa & ";true"
    aElemento(2, 3) = "NombreAfp;" & Trim(lblHelp(0)) & ";true"
    aElemento(2, 4) = "SinPago;" & sSinPago & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "codcls char(2) Not Null, codpsn varchar(11) Not Null, numeroafp varchar(15) Null, apepaterno varchar(25) Null, "
    s_Sql = s_Sql & "apematerno varchar(25) Null, nombres varchar(25) Null, fecingregpen date Null, "
    s_Sql = s_Sql & "direccionpsn varchar(200) Null, telefonopsn varchar(20) Null, "
    s_Sql = s_Sql & "estadopsn char(1) Null, fecestado date Null, direccion varchar(50) Null, "
    s_Sql = s_Sql & "numero varchar(15) Null, localidad varchar(50) Null, departamento varchar(25) Null, "
    s_Sql = s_Sql & "provincia varchar(25) Null, distrito varchar(25) Null, postal char(2) Null, "
    s_Sql = s_Sql & "telefono varchar(20) Null, representante varchar(80) Null, tipodocu char(2) Null, "
    s_Sql = s_Sql & "documento varchar(15) Null, encargado varchar(80) Null, email varchar(40) Null, "
    s_Sql = s_Sql & "cencosto varchar(40) Null, anexo varchar(20) Null, moneda char(1) Null, "
    s_Sql = s_Sql & "nrohoja varchar(10) Null, nroafiliados int(4) DEFAULT '0', fechapago date Null, formapago char(1) Null, "
    s_Sql = s_Sql & "interespension decimal(18,2) Null, chequepension varchar(10) Null, bcopension varchar(40) Null, "
    s_Sql = s_Sql & "interesadmin decimal(18,2) Null, chequeadmin varchar(10) Null, bcoadmin varchar(40) Null, "
    s_Sql = s_Sql & "impremase decimal(18,2) Null Default '0', impapobli decimal(18,2) Null Default '0',"
    s_Sql = s_Sql & "impapovolsfp decimal(18,2) Null Default '0', impapovolcfp decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "impapoemp decimal(18,2) Null Default '0', impseguro decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "impporcen decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "ejercicio varchar(4) Null, "
    s_Sql = s_Sql & "periodo char(2) Null, "
    s_Sql = s_Sql & "afiliado char(1) DEFAULT '0', "
    s_Sql = s_Sql & "imp27252 decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "firma blob NULL , "
    s_Sql = s_Sql & "PRIMARY KEY (codcls, codpsn)) "
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Genera la información del reporte
    PlanillaPrevision "tmp" & gdl_Procedure.ps_ReportName, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "R"
    
    If chkfirma.Value = Checked Then
    
        s_Sql = "UPDATE tmp" & gdl_Procedure.ps_ReportName & ", plcfgempresa "
        s_Sql = s_Sql & "SET tmp" & gdl_Procedure.ps_ReportName & ".firma=plcfgempresa.firma "
        s_Sql = s_Sql & "Where plcfgempresa.pdoano='" & ps_Anyo & "'"
        Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    
    End If
    
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 7), False, True, False, True, True, aElemento, aElementos, porstRecordset, "rptDesafiliados"
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_SwRegistro, "", ps_Usuario, s_FechaHora, "E"
  End Select

End Sub


Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Entiodad de fondo de pensiones(afp)
    tdbHelp.Columns(0).DataField = "codafp": tdbHelp.Columns(1).DataField = "desafp"
    tdbHelp.Caption = "Entidad Pensiones"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("afp", "codafp", "", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = frmCuadro.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = 1470
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

    dir.path = "C:\"
    
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
  cnn.CursorLocation = adUseClient
  cnn.Open
  

  ' Establece posición del formulario
  Me.Height = 7470: Me.Width = 6230
  Me.Left = 520: Me.Top = 300
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Planilla Previsional de AFP"
  s_TitleTable = "Clase Planilla"
  
  ReDim aElemento(3, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcls", "descls", "estadocls")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 800, 3556.03, 300)
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
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(2).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(2).ValueItems.Translate = True
  For n_Index = 0 To 1
    tdbRegistro.Columns(2).ValueItems.Add Item
    tdbRegistro.Columns(2).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Act, s_Estado_Ina)
    tdbRegistro.Columns(2).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadnok") & ".bmp")
  Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbRegistro.AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(9, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "promedio", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Parametros", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
 '[ Configuración el control de ayuda
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  For n_Index = 1 To 12: cmbMes.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codafp", "desafp")
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
  
  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
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
   Case 0       ' Entidad de pensiones (afp)
    txtAfp = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtAfp.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Entidad de fondo de pensiones(afp)
    s_Sql = gdl_Funcion.HelpTablas("afp", tdbHelp.Columns(ColIndex).DataField, "", "")
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
Private Sub txtAfp_GotFocus()
  gdl_Procedure.MarcaGet txtAfp
End Sub
Private Sub txtAfp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtAfp_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub txtAfp_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtAfp, "EP")
End Sub

Private Sub Dir_Change()
If Len(dir.path) = 3 Then
  ruta.Caption = dir.path
Else
  ruta.Caption = dir.path & "\"
End If
End Sub

Private Sub Drive_Change()
On Error GoTo Err_
dir.path = drive.drive
Exit Sub
Err_: drive.drive = "C:\"
End Sub

Private Sub cmdbusqueda_Click()
    dir.Refresh
    If cmdbusqueda.Outline = False Then
    dir.Visible = True
    cmdbusqueda.Outline = True
    Else
    dir.Visible = False
    cmdbusqueda.Outline = False
    End If
End Sub

Private Sub cmdexportar_Click()


    Dim Rst As ADODB.Recordset
    Dim R As Boolean
    Dim sql As String
    Dim AnioMes As String
        
    AnioMes = ps_Anyo & LTrim(Left(cmbMes.Text, 2))
        
        
    If cmbMes.Text = "" Then MsgBox " Falta seleccionar el Mes", vbInformation: Exit Sub
    mesanterior = Int(Left(cmbMes.Text, 2)) - 1
    
    If Len(mesanterior) = 1 Then
      mesanterior = "0" & mesanterior
    End If
    
    Set Rst = New ADODB.Recordset
         
'    '***sql = sql & " if(month(psn.fecestado) in ('" & mesanterior & "','" & Left(cmbmes.Text, 2) & "') and year(psn.fecestado)='" & ps_Anyo & "' ,"
'    sql = "SELECT ifnull(psn.numeroafp,''), (CASE LEFT(dci.codsunat, 2) WHEN '01' THEN '0' WHEN '04' THEN '1' WHEN '07' THEN '4' WHEN '11' THEN '3' ELSE '5' END), "
'    sql = sql & "psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres, "
'    sql = sql & "IF( psn.estadopsn <> 'A' , CASE psn.estadopsn "
'    sql = sql & "WHEN 'V' THEN '5' WHEN 'L' THEN '4' WHEN 'N' THEN '3' WHEN 'P' THEN '3' WHEN 'I' THEN '2' END, "
'    sql = sql & "IF(month(psn.fecestado)='" & Left(cmbmes.Text, 2) & "' and year(psn.fecestado)='" & ps_Anyo & "','1','') ) as cod, "
'    '***sql = sql & " if(month(psn.fecestado) in ('" & mesanterior & "','" & Left(cmbmes.Text, 2) & "') and year(psn.fecestado)='" & ps_Anyo & "' ,date_format(psn.fecestado,'%d/%m/%Y'),''),"
'    sql = sql & "IF( psn.estadopsn <> 'A', date_format(psn.fecestado,'%d/%m/%Y'), "
'    sql = sql & "IF(month(psn.fecestado)='" & Left(cmbmes.Text, 2) & "' and year(psn.fecestado)='" & ps_Anyo & "', date_format(psn.fecestado,'%d/%m/%Y'),'') ) AS fec, "
'    sql = sql & "ROUND(SUM(importe_mn), 2), '' AS I1, '' AS I2, '' AS I3, 'N' as rubro1, '' AS rubro2 "
'    sql = sql & "FROM plresultado res "
'    sql = sql & "INNER JOIN plpersonal psn on res.codcls=psn.codcls and res.codpsn=psn.codpsn "
'    sql = sql & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
'    sql = sql & "INNER JOIN plentidadafp afp on psn.codafp=afp.codafp "
'    sql = sql & "INNER JOIN plparametroafp pafp on res.pdoano=pafp.pdoano "
'    sql = sql & "WHERE res.pdoano='" & ps_Anyo & "' "
'    sql = sql & "AND res.pdomes='" & Left(cmbmes.Text, 2) & "' "
'    sql = sql & "AND res.codcpc=pafp.cpcremase "
'    sql = sql & "AND left(afp.codsunat,2) IN ('21','22','23','24', '25') "
'    sql = sql & "GROUP BY res.codcls,res.codpsn,res.codcpc "
'    sql = sql & "ORDER BY psn.apepaterno,psn.apematerno,psn.nombres"


    '** MODIFICACION AGOSTO 2015, Nuevo Formato AFPNET
     
     sql = "SELECT ifnull(psn.numeroafp,''), (CASE LEFT(dci.codsunat, 2) WHEN '01' THEN '0' WHEN '04' THEN '1' WHEN '07' THEN '4' WHEN '11' THEN '3' ELSE '5' END), psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres, "
     sql = sql & "CASE WHEN estadopsn='I' AND date_format(fecbaja,'%Y%m')='" & AnioMes & "' OR fecbaja IS NULL THEN 'S' ELSE 'N' end as rl_laboral, "
     sql = sql & "CASE WHEN date_format(fecingreso,'%Y%m')='" & AnioMes & "' THEN 'S' ELSE 'N' END AS ini_rlaboral, "
     sql = sql & "CASE WHEN estadopsn='I' AND date_format(fecbaja,'%Y%m')='" & AnioMes & "' THEN 'S' ELSE 'N' end as cese_rl, "
     sql = sql & "CASE WHEN xasis.codmdi_licen='05' THEN 'L' ELSE ' ' END as exepcion_aporte, "
     sql = sql & "ROUND(SUM(importe_mn), 2), '' AS I1, '' AS I2, '' AS I3, 'N' as rubro1, '' AS rubro2 "
     sql = sql & "FROM plresultado res "
     sql = sql & "INNER JOIN plpersonal psn on res.codcls=psn.codcls and res.codpsn=psn.codpsn "
     sql = sql & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
     sql = sql & "INNER JOIN plentidadafp afp on psn.codafp=afp.codafp "
     sql = sql & "INNER JOIN plparametroafp pafp on res.pdoano=pafp.pdoano "
     sql = sql & "INNER JOIN plasistencia xasis ON xasis.codcls=res.codcls AND xasis.codpdo=res.codpdo and xasis.codpsn=res.codpsn "
     sql = sql & "WHERE res.pdoano='" & ps_Anyo & "' "
     sql = sql & "AND res.pdomes='" & Left(cmbMes.Text, 2) & "' "
     sql = sql & "AND res.codcpc=pafp.cpcremase "
     sql = sql & "AND left(afp.codsunat,2) IN ('21','22','23','24', '25') "
     sql = sql & "GROUP BY res.codcls,res.codpsn,res.codcpc "
     sql = sql & "ORDER BY psn.apepaterno,psn.apematerno,psn.nombres"
     Rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    
    If Rst.RecordCount = 0 Then
      MsgBox " No Existen datos para Generar el Archivo ", vbInformation
      Exit Sub
    End If
    
    R = Recordset_a_Csv(Rst, ruta.Caption & ps_RucEmpresa & Format(Day(Date), "00") & Format(Month(Date), "00") & Year(Date) & ".TXT")
        
    MsgBox " Se generó el archivo " & ps_RucEmpresa & Format(Day(Date), "00") & Format(Month(Date), "00") & Year(Date) & ".TXT" & " correctamente en base a " & Rst.RecordCount & " Registros", vbInformation
    
    If Not Rst.State = adStateOpen Then
        Rst.Close
    End If
    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

End Sub

Function Recordset_a_Csv_ORIGINal(rs As Recordset, path As String) As Boolean
  Dim cadenav As String, ceros As String
  Dim Fila As Integer, desde As Integer
  Dim tamanno As Integer, posicion As Integer
  Dim Columna As Long
  
  On Error GoTo Err_function
  ' Crea el archivo
  Open path For Output As #1
  ' Se mueve al primer registro
  rs.MoveFirst
  
  ' recorre todo el recordset
  For Fila = 0 To rs.RecordCount - 1
    ' nombre del campo
    cadenav = ""
    tamanno = 12
    valor = Trim(rs.Fields(0))
    
    If Len(valor) = tamanno Then
    ElseIf Len(valor) > tamanno Then
      valor = Left(valor, tamanno)
    ElseIf Len(valor) < tamanno Then
      cadenav = gdl_Funcion.PadR(cadenav, (tamanno - Len(valor)), " ")
    End If
    valor = valor & cadenav
  
    Print #1, Format(Fila + 1, "00000") & valor;
    ' recorre todos los campos
    For Columna = 1 To rs.Fields.Count - 1
      ' imprime la fila actual en el fichero
      valor = IIf(IsNull(rs.Fields(Columna)), "", Trim(rs.Fields(Columna)))
      cadenav = ""
      
      Select Case Columna
       Case 1, 6, 12
        tamanno = 1
       Case 2, 3, 4, 5
        tamanno = 20
       Case 7
        tamanno = 10
       Case 8, 9, 10, 11
        tamanno = 9
       Case 13
        tamanno = 2
      End Select
    
      If Len(valor) = tamanno Then
      ElseIf Len(valor) > tamanno Then
        valor = Left(valor, tamanno)
      ElseIf Len(valor) < tamanno Then
        If Columna <> 8 Then
          cadenav = gdl_Funcion.PadR(cadenav, (tamanno - Len(valor)), " ")
        End If
      End If
    
      If Columna = 8 Then
        posicion = InStr(valor, ".")
        If Len(valor) - posicion = 1 Then
          valor = valor & "0"
        End If
        If posicion = 0 Then
          valor = valor & ".00"
        End If
        posicion = InStr(valor, ".")
        While 6 > posicion - 1
          ceros = "0" & ceros
          posicion = posicion + 1
        Wend
        valor = ceros & valor
        ceros = ""
      End If
    
      If (Columna >= 9 And Columna <= 11) Then
        cadenav = "000000.00"
      End If
      valor = valor & cadenav
      Print #1, "" & valor;
    Next
    ' escribe una línea en blanco
    Print ""
    ' salto de carro
    Print #1, "" & Chr(13) & Chr(10);
    ' mueve el recordset al siguiente registro
    rs.MoveNext
  Next
  ' cierra el archivo
  Close #1
  Exit Function

Err_function:
  MsgBox Err.Description, vbCritical
  Close
End Function

Private Sub cmdexcel_Click()
    Dim Rst As ADODB.Recordset
    Dim R As Boolean
    Dim sql As String
    Dim AnioMes As String
        
  AnioMes = ps_Anyo & LTrim(Left(cmbMes.Text, 2))
  
  If cmbMes.Text = "" Then MsgBox " Falta seleccionar el Mes", vbInformation: Exit Sub
  mesanterior = Int(Left(cmbMes.Text, 2)) - 1
  
  If Len(mesanterior) = 1 Then
    mesanterior = "0" & mesanterior
  End If
    
  Set Rst = New ADODB.Recordset
'  sql = "SELECT 'x' as cont, psn.numeroafp, (CASE LEFT(dci.codsunat, 2) WHEN '01' THEN '0' WHEN '04' THEN '1' WHEN '07' THEN '4' WHEN '11' THEN '3' ELSE '5' END) AS dcisunat, psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres, "
'  sql = sql & "IF( psn.estadopsn <> 'A' , CASE psn.estadopsn "
'  sql = sql & "WHEN 'V' THEN '5' WHEN 'L' THEN '4' WHEN 'N' THEN '3' WHEN 'P' THEN '3' WHEN 'I' THEN '2' END, "
'  sql = sql & "IF(month(psn.fecestado)='" & Left(cmbmes.Text, 2) & "' AND year(psn.fecestado)='" & ps_Anyo & "','1','')) AS cod, "
'  'sql = sql & " if(month(psn.fecestado) in ('" & mesanterior & "','" & Left(cmbmes.Text, 2) & "') and year(psn.fecestado)='" & ps_Anyo & "' ,date_format(psn.fecestado,'%d/%m/%Y'),''),"
'  sql = sql & "IF( psn.estadopsn <> 'A' , date_format(psn.fecestado,'%d/%m/%Y'), "
'  sql = sql & "IF(month(psn.fecestado)='" & Left(cmbmes.Text, 2) & "' AND year(psn.fecestado)='" & ps_Anyo & "', date_format(psn.fecestado,'%d/%m/%Y'),'') ) AS fec, "
'  sql = sql & "SUM(importe_mn),'' as I1,'' as I2,'' as I3, 'N' as rubro1,'' as rubro2 "
'  sql = sql & "FROM plresultado res "
'  sql = sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
'  sql = sql & "INNER JOIN plentidadafp afp ON psn.codafp=afp.codafp "
'  sql = sql & "INNER JOIN plparametroafp pafp ON res.pdoano=pafp.pdoano "
'  sql = sql & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
'  sql = sql & "WHERE res.pdoano='" & ps_Anyo & "' AND res.pdomes='" & Left(cmbmes.Text, 2) & "' AND res.codcpc=pafp.cpcremase "
'  sql = sql & "AND left(afp.codsunat,2) in ('21','22','23','24', '25') "
'  sql = sql & "GROUP BY res.codcls,res.codpsn,res.codcpc "
'  sql = sql & "ORDER BY psn.apepaterno,psn.apematerno,psn.nombres"

'**MODIFICACION AGOSTO Y SEPTIEMBRE 2015
    sql = "SELECT ifnull(psn.numeroafp,''), (CASE LEFT(dci.codsunat, 2) WHEN '01' THEN '0' WHEN '04' THEN '1' WHEN '07' THEN '4' WHEN '11' THEN '3' ELSE '5' END), psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres, "
    sql = sql & "case when estadopsn='I' AND date_format(fecbaja,'%Y%m')='" & AnioMes & "' OR fecbaja IS Null THEN 'S' ELSE 'N' end as rl_laboral, "
    sql = sql & "CASE WHEN date_format(fecingreso,'%Y%m')='" & AnioMes & "' THEN 'S' ELSE 'N' END AS ini_rlaboral, "
    sql = sql & "CASE WHEN estadopsn='I' AND date_format(fecbaja,'%Y%m')='" & AnioMes & "' THEN 'S' ELSE 'N' end as cese_rl, "
    sql = sql & "CASE WHEN xasis.codmdi_licen='05' THEN 'L' ELSE ' ' END as exepcion_aporte, "
    sql = sql & "ROUND(SUM(importe_mn), 2), '' AS I1, '' AS I2, "
    
    sql = sql & "(SELECT IFNULL(ROUND(SUM(resx.importe_mn), 2),0) FROM plresultado resx where resx.codcls=res.codcls AND resx.codpdo=res.codpdo AND res.pdoano=resx.pdoano AND resx.codcpc=pafp.cpcapoemp AND resx.codpsn=psn.codpsn) AS I3, "
    
    sql = sql & "'N' as rubro1, '' AS rubro2 "
    sql = sql & "FROM plresultado res INNER JOIN plpersonal psn on res.codcls=psn.codcls and res.codpsn=psn.codpsn INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci INNER JOIN plentidadafp afp on psn.codafp=afp.codafp INNER JOIN plparametroafp pafp on res.pdoano=pafp.pdoano "
    sql = sql & "INNER JOIN plasistencia xasis ON xasis.codcls=res.codcls AND xasis.codpdo=res.codpdo and xasis.codpsn=res.codpsn "
    sql = sql & "WHERE res.pdoano='" & ps_Anyo & "' "
    sql = sql & "AND res.pdomes='" & Left(cmbMes.Text, 2) & "' "
    sql = sql & "AND res.codcpc=pafp.cpcremase AND left(afp.codsunat,2) IN ('21','22','23','24', '25') "
    sql = sql & "GROUP BY res.codcls,res.codpsn,res.codcpc "
    sql = sql & "ORDER BY psn.apepaterno,psn.apematerno,psn.nombres"
  Rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    
  If Rst.RecordCount = 0 Then
    MsgBox " No Existen datos para Generar el Archivo ", vbInformation
    Exit Sub
  End If
    
Dim nhoja As String
Dim i As Integer
Dim j As Integer
Dim cols As Integer
nhoja = "AFPNET"
cols = 16
Dim ApExcel As Variant
Set ApExcel = CreateObject("Excel.application")
ApExcel.Visible = False
ApExcel.Workbooks.Add
ApExcel.Sheets("Hoja1").Name = nhoja
ApExcel.ActiveWindow.Zoom = 75
'************************************
On Error GoTo Error
  Rst.MoveFirst
  For i = 1 To Rst.RecordCount
     ApExcel.Cells(i, 1).Formula = i
     ApExcel.Cells(i, 2).Formula = Rst(0).Value
     ApExcel.Cells(i, 3).Formula = Rst(1).Value
    
    
     For j = 3 To 16 'cols
      
        If j = 11 Then
          ApExcel.Cells(i, j + 1).Formula = Round(Rst(j - 1), 2)
          ApExcel.Cells(i, j + 1).NumberFormat = "0.00"
        Else
          ApExcel.Cells(i, j + 1).NumberFormat = "@"
          ApExcel.Cells(i, j + 1).Formula = Rst(j - 1).Value
          
        End If
   
    Next
    Rst.MoveNext
  Next
    
  If Not Rst.State = adStateOpen Then
    Rst.Close
  End If
  If Not Rst Is Nothing Then
    Set Rst = Nothing
  End If
  TipodeProgreso = 1
  IntervalodeTiempo = 100
  labelprogreso = "Generando                                                          Excel"
  Progreso.Show vbModal
  MsgBox ("Proceso de Exportacion a Excel, terminado")
  ApExcel.Visible = True
Error:
If Err.Number <> 0 Then
   MsgBox Err.Description, vbCritical, ps_NomSistema
   Exit Sub
End If
End Sub
Private Sub cmdverificar_Click()
  Dim Rst As ADODB.Recordset
  Dim R As Boolean
  Dim sql As String
  
  Set Rst = New ADODB.Recordset
    
  sql = "SELECT (CASE LEFT(dci.codsunat, 2) WHEN '01' THEN '0' WHEN '04' THEN '1' WHEN '07' THEN '4' WHEN '11' THEN '3' ELSE '5' END), psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres "
  sql = sql & "FROM plpersonal psn "
  sql = sql & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  sql = sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' AND psn.estadopsn='A' "
  Rst.Open sql, cnn, adOpenStatic, adLockOptimistic
  
  If Rst.RecordCount = 0 Then
    MsgBox " No Existen datos para Generar el Archivo ", vbInformation
    Exit Sub
  End If
    
  Dim nhoja As String
  Dim i As Integer
  Dim j As Integer
  Dim cols As Integer
  nhoja = "RUC " & ps_RucEmpresa
  cols = 5
  
  Dim ApExcel As Variant
  Set ApExcel = CreateObject("Excel.application")
  ApExcel.Visible = False
  ApExcel.Workbooks.Add
  ApExcel.Sheets("Hoja1").Name = nhoja
  ApExcel.ActiveWindow.Zoom = 75
  
  '************************************
  On Error GoTo Error
  Rst.MoveFirst
  For i = 1 To Rst.RecordCount
    For j = 1 To cols
      ApExcel.Cells(i, j).Formula = Rst(j - 1)
    Next
    Rst.MoveNext
  Next
  TipodeProgreso = 1
  IntervalodeTiempo = 100
  labelprogreso = "Generando                                                          Excel"
  Progreso.Show vbModal
  MsgBox ("Proceso de Exportacion a Excel, terminado")
  ApExcel.Visible = True
Error:

End Sub
