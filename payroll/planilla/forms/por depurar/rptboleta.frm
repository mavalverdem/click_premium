VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fReporteBoleta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "rptboleta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   7905
   Begin Threed.SSPanel panToolBar 
      Height          =   5625
      Index           =   0
      Left            =   7125
      TabIndex        =   7
      Top             =   975
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9922
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
         TabIndex        =   19
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
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1245
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
         Picture         =   "rptboleta.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   1665
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
         Picture         =   "rptboleta.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   2370
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
         Picture         =   "rptboleta.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   2805
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
         Picture         =   "rptboleta.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   14
         Tag             =   "0"
         Top             =   3915
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
         Picture         =   "rptboleta.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   15
         Tag             =   "0"
         Top             =   4350
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
         Picture         =   "rptboleta.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   810
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
         Picture         =   "rptboleta.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   13
         Tag             =   "0"
         Top             =   3225
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
         Picture         =   "rptboleta.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   16
         Tag             =   "0"
         Top             =   4785
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
         Picture         =   "rptboleta.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   9
         Left            =   240
         TabIndex        =   17
         Tag             =   "1"
         Top             =   810
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
         Picture         =   "rptboleta.frx":0108
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   285
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Otros"
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
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Tag             =   "1"
         Top             =   1245
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
         Picture         =   "rptboleta.frx":0124
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
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
      Begin VB.TextBox txtPeriodo 
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   1455
         TabIndex        =   5
         Top             =   495
         Width           =   810
      End
      Begin VB.TextBox txtFormato 
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   1455
         TabIndex        =   2
         Top             =   150
         Width           =   810
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   2325
         TabIndex        =   23
         Top             =   150
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
         Left            =   2325
         TabIndex        =   24
         Top             =   495
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSRibbon ribCopia 
         Height          =   360
         Left            =   6855
         TabIndex        =   33
         Top             =   510
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   3
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "rptboleta.frx":0140
      End
      Begin Threed.SSRibbon ribFirma 
         Height          =   360
         Left            =   7260
         TabIndex        =   34
         Top             =   510
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   4
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "rptboleta.frx":015C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   6855
         TabIndex        =   30
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
         PictureUp       =   "rptboleta.frx":0178
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   6450
         TabIndex        =   29
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
         PictureUp       =   "rptboleta.frx":0194
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   7260
         TabIndex        =   31
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
         PictureUp       =   "rptboleta.frx":01B0
      End
      Begin Threed.SSRibbon ribOrdenar 
         Height          =   360
         Left            =   6450
         TabIndex        =   32
         Top             =   510
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "rptboleta.frx":01CC
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
         Left            =   2685
         TabIndex        =   3
         Top             =   195
         Width           =   180
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
         Index           =   1
         Left            =   2685
         TabIndex        =   6
         Top             =   540
         Width           =   180
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Formato :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   195
         Width           =   1005
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   0
         Left            =   270
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   5895
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   5625
      Left            =   45
      TabIndex        =   21
      Top             =   975
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9922
      _Version        =   393216
      TabOrientation  =   1
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
      TabCaption(0)   =   "Personal"
      TabPicture(0)   =   "rptboleta.frx":01E8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dcaSeleccion(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbSeleccion(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Centro Costo"
      TabPicture(1)   =   "rptboleta.frx":0204
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdbSeleccion(1)"
      Tab(1).Control(1)=   "dcaSeleccion(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ubicación"
      TabPicture(2)   =   "rptboleta.frx":0220
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tdbSeleccion(2)"
      Tab(2).Control(1)=   "dcaSeleccion(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Sección"
      TabPicture(3)   =   "rptboleta.frx":023C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tdbSeleccion(3)"
      Tab(3).Control(1)=   "dcaSeleccion(3)"
      Tab(3).ControlCount=   2
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4365
         Index           =   1
         Left            =   -74280
         TabIndex        =   25
         Top             =   90
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   7699
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
         Left            =   -74280
         Top             =   4455
         Width           =   5460
         _ExtentX        =   9631
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
         Height          =   4360
         Index           =   2
         Left            =   -74280
         TabIndex        =   26
         Top             =   90
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   7699
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
         Left            =   -74280
         Top             =   4455
         Width           =   5460
         _ExtentX        =   9631
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
         Height          =   4365
         Index           =   3
         Left            =   -74280
         TabIndex        =   27
         Top             =   90
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   7699
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
         Left            =   -74280
         Top             =   4455
         Width           =   5460
         _ExtentX        =   9631
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
         Height          =   4845
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   60
         Width           =   6840
         _ExtentX        =   12065
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
      Begin MSAdodcLib.Adodc dcaSeleccion 
         Height          =   330
         Index           =   0
         Left            =   90
         Top             =   4935
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
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2835
      TabIndex        =   28
      Top             =   720
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
Attribute VB_Name = "fReporteBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(4, 2)                            ' Array de inicio y fin de seleccion de registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Dim s_FmtBoleta As String
    
Dim fso As Object                                       'ENE 2015 Variable de objeto para usar en funcion buscar un archivo

'[
Private Sub Imprime_Detalle(sFormato As String, sPersona As String, sPeriodo As String, sFont As String, nLinea2ndFmt As Integer)
  
  Dim rsIngresos As New ADODB.Recordset
  Dim rsDesctos As New ADODB.Recordset
  Dim rsAportes As New ADODB.Recordset
  Dim rsFormato As New ADODB.Recordset

  Dim sSQL As String
  Dim lProcesa As Boolean
  Dim nFila As Integer

  Dim lIngresos As Boolean
  Dim lDesctos As Boolean
  Dim lAportes As Boolean

  Dim lRsIngresos As Boolean
  Dim lRsDesctos As Boolean
  Dim lRsAportes As Boolean

  lIngresos = False
  lDesctos = False
  lAportes = False
  
  lRsIngresos = False
  lRsDesctos = False
  lRsAportes = False
  
  Dim nIngresos_CodigoX As Integer
  Dim nIngresos_CodigoY As Integer
  Dim nIngresos_DescripcionX As Integer
  Dim nIngresos_DescripcionY As Integer
  Dim nIngresos_ImporteX As Integer
  Dim nIngresos_ImporteY As Integer
  
  Dim nDesctos_CodigoX As Integer
  Dim nDesctos_CodigoY As Integer
  Dim nDesctos_DescripcionX As Integer
  Dim nDesctos_DescripcionY As Integer
  Dim nDesctos_ImporteX As Integer
  Dim nDesctos_ImporteY As Integer
  
  Dim nAportes_CodigoX As Integer
  Dim nAportes_CodigoY As Integer
  Dim nAportes_DescripcionX As Integer
  Dim nAportes_DescripcionY As Integer
  Dim nAportes_ImporteX As Integer
  Dim nAportes_ImporteY As Integer
  
  Dim sIngresosC_FSize As String
  Dim sIngresosC_Font As String
  Dim sIngresosC_FontN As String
  Dim sIngresosC_FontS As String
  Dim sIngresosC_FontC As String
  Dim sIngresosC_FontTipo As String
  Dim nIngresosC_Longitud As Integer
  
  Dim sIngresosD_FSize As String
  Dim sIngresosD_Font As String
  Dim sIngresosD_FontN As String
  Dim sIngresosD_FontS As String
  Dim sIngresosD_FontC As String
  Dim sIngresosD_FontTipo As String
  Dim nIngresosD_Longitud As Integer
  
  Dim sIngresosI_FSize As String
  Dim sIngresosI_Font As String
  Dim sIngresosI_FontN As String
  Dim sIngresosI_FontS As String
  Dim sIngresosI_FontC As String
  Dim sIngresosI_FontTipo As String
  Dim nIngresosI_Longitud As Integer
  
  Dim sDesctosC_FSize As String
  Dim sDesctosC_Font As String
  Dim sDesctosC_FontN As String
  Dim sDesctosC_FontS As String
  Dim sDesctosC_FontC As String
  Dim sDesctosC_FontTipo As String
  Dim nDesctosC_Longitud As Integer
  
  Dim sDesctosD_FSize As String
  Dim sDesctosD_Font As String
  Dim sDesctosD_FontN As String
  Dim sDesctosD_FontS As String
  Dim sDesctosD_FontC As String
  Dim sDesctosD_FontTipo As String
  Dim nDesctosD_Longitud As Integer
  
  Dim sDesctosI_FSize As String
  Dim sDesctosI_Font As String
  Dim sDesctosI_FontN As String
  Dim sDesctosI_FontS As String
  Dim sDesctosI_FontC As String
  Dim sDesctosI_FontTipo As String
  Dim nDesctosI_Longitud As Integer
  
  Dim sAportesC_FSize As String
  Dim sAportesC_Font As String
  Dim sAportesC_FontN As String
  Dim sAportesC_FontS As String
  Dim sAportesC_FontC As String
  Dim sAportesC_FontTipo As String
  Dim nAportesC_Longitud As Integer
  
  Dim sAportesD_FSize As String
  Dim sAportesD_Font As String
  Dim sAportesD_FontN As String
  Dim sAportesD_FontS As String
  Dim sAportesD_FontC As String
  Dim sAportesD_FontTipo As String
  Dim nAportesD_Longitud As Integer
  
  Dim sAportesI_FSize As String
  Dim sAportesI_Font As String
  Dim sAportesI_FontN As String
  Dim sAportesI_FontS As String
  Dim sAportesI_FontC As String
  Dim sAportesI_FontTipo As String
  Dim nAportesI_Longitud As Integer
  
  'Inicializa Variables
  nFila = 0
  nIngresos_CodigoX = 0: nIngresos_CodigoY = 0: nIngresos_DescripcionX = 0: nIngresos_DescripcionY = 0: nIngresos_ImporteX = 0: nIngresos_ImporteY = 0
  nDesctos_CodigoX = 0: nDesctos_CodigoY = 0: nDesctos_DescripcionX = 0: nDesctos_DescripcionY = 0: nDesctos_ImporteX = 0: nDesctos_ImporteY = 0
  nAportes_CodigoX = 0: nAportes_CodigoY = 0: nAportes_DescripcionX = 0: nAportes_DescripcionY = 0: nAportes_ImporteX = 0: nAportes_ImporteY = 0
  
  'Datos de Detalle
  sSQL = "SELECT seccion, dato, tipodato, nombre, fila, columna, longitud, sizefont, fontn, fonts, fontc "
  sSQL = sSQL & "FROM pldetaboleta fbl "
  sSQL = sSQL & "INNER JOIN plvarfunc pvf ON pvf.codigo = fbl.dato "
  sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND codboleta='" & sFormato & "' "
  sSQL = sSQL & "AND seccion='D' "
  sSQL = sSQL & "ORDER BY fila, columna"
  Set rsFormato = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (rsFormato.EOF And rsFormato.BOF) Then
    nFila = CInt(rsFormato("fila"))
    Do While Not rsFormato.EOF
      If CInt(nFila) > CInt(rsFormato("fila")) Then
        nFila = rsFormato("fila")
      End If
      Select Case rsFormato("nombre")
        Case "Ingresos_Codigo": nIngresos_CodigoX = rsFormato("columna"): nIngresos_CodigoY = rsFormato("fila"): lIngresos = True: sIngresosC_FSize = rsFormato("sizefont"): sIngresosC_Font = sFont: sIngresosC_FontN = rsFormato("fontn"): sIngresosC_FontS = rsFormato("fonts"): sIngresosC_FontC = rsFormato("fontc"): sIngresosC_FontTipo = rsFormato("tipodato"): nIngresosC_Longitud = rsFormato("Longitud")
        Case "Ingresos_Descripcion": nIngresos_DescripcionX = rsFormato("columna"): nIngresos_DescripcionY = rsFormato("fila"): lIngresos = True: sIngresosD_FSize = rsFormato("sizefont"): sIngresosD_Font = sFont: sIngresosD_FontN = rsFormato("fontn"): sIngresosD_FontS = rsFormato("fonts"): sIngresosD_FontC = rsFormato("fontc"): sIngresosD_FontTipo = rsFormato("tipodato"): nIngresosD_Longitud = rsFormato("Longitud")
        Case "Ingresos_Importe": nIngresos_ImporteX = rsFormato("columna"): nIngresos_ImporteY = rsFormato("fila"): lIngresos = True: sIngresosI_FSize = rsFormato("sizefont"): sIngresosI_Font = sFont: sIngresosI_FontN = rsFormato("fontn"): sIngresosI_FontS = rsFormato("fonts"): sIngresosI_FontC = rsFormato("fontc"): sIngresosI_FontTipo = rsFormato("tipodato"): nIngresosI_Longitud = rsFormato("Longitud")
       
        Case "Desctos_Codigo": nDesctos_CodigoX = rsFormato("columna"): nDesctos_CodigoY = rsFormato("fila"): lDesctos = True: sDesctosC_FSize = rsFormato("sizefont"): sDesctosC_Font = sFont: sDesctosC_FontN = rsFormato("fontn"): sDesctosC_FontS = rsFormato("fonts"): sDesctosC_FontC = rsFormato("fontc"): sDesctosC_FontTipo = rsFormato("tipodato"): nDesctosC_Longitud = rsFormato("Longitud")
        Case "Desctos_Descripcion": nDesctos_DescripcionX = rsFormato("columna"): nDesctos_DescripcionY = rsFormato("fila"): lDesctos = True: sDesctosD_FSize = rsFormato("sizefont"): sDesctosD_Font = sFont: sDesctosD_FontN = rsFormato("fontn"): sDesctosD_FontS = rsFormato("fonts"): sDesctosD_FontC = rsFormato("fontc"): sDesctosD_FontTipo = rsFormato("tipodato"): nDesctosD_Longitud = rsFormato("Longitud")
        Case "Desctos_Importe": nDesctos_ImporteX = rsFormato("columna"): nDesctos_ImporteY = rsFormato("fila"): lDesctos = True: sDesctosI_FSize = rsFormato("sizefont"): sDesctosI_Font = sFont: sDesctosI_FontN = rsFormato("fontn"): sDesctosI_FontS = rsFormato("fonts"): sDesctosI_FontC = rsFormato("fontc"): sDesctosI_FontTipo = rsFormato("tipodato"): nDesctosI_Longitud = rsFormato("Longitud")
      
        Case "Aportes_Codigo": nAportes_CodigoX = rsFormato("columna"): nAportes_CodigoY = rsFormato("fila"): lAportes = True: sAportesC_FSize = rsFormato("sizefont"): sAportesC_Font = sFont: sAportesC_FontN = rsFormato("fontn"): sAportesC_FontS = rsFormato("fonts"): sAportesC_FontC = rsFormato("fontc"): sAportesC_FontTipo = rsFormato("tipodato"): nAportesC_Longitud = rsFormato("Longitud")
        Case "Aportes_Descripcion": nAportes_DescripcionX = rsFormato("columna"): nAportes_DescripcionY = rsFormato("fila"): lAportes = True: sAportesD_FSize = rsFormato("sizefont"): sAportesD_Font = sFont: sAportesD_FontN = rsFormato("fontn"): sAportesD_FontS = rsFormato("fonts"): sAportesD_FontC = rsFormato("fontc"): sAportesD_FontTipo = rsFormato("tipodato"): nAportesD_Longitud = rsFormato("Longitud")
        Case "Aportes_Importe": nAportes_ImporteX = rsFormato("columna"): nAportes_ImporteY = rsFormato("fila"): lAportes = True: sAportesI_FSize = rsFormato("sizefont"): sAportesI_Font = sFont: sAportesI_FontN = rsFormato("fontn"): sAportesI_FontS = rsFormato("fonts"): sAportesI_FontC = rsFormato("fontc"): sAportesI_FontTipo = rsFormato("tipodato"): nAportesI_Longitud = rsFormato("Longitud")
      End Select
      rsFormato.MoveNext
    Loop
    rsFormato.Close
  End If
  
  'Selecciona Conceptos de Ingresos
  If lIngresos Then
    sSQL = "SELECT res.codcpc, cpc.descpc, res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS importe "
    sSQL = sSQL & "FROM plresultado res "
    sSQL = sSQL & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    sSQL = sSQL & "AND res.codpdo='" & sPeriodo & "' "
    sSQL = sSQL & "AND res.codpsn='" & sPersona & "' "
    sSQL = sSQL & "AND res.tipocpc='" & s_Estado_Ina & "' "
    sSQL = sSQL & "AND res.impbolecpc='" & s_Estado_Act & "' "
    sSQL = sSQL & "ORDER BY res.secuencia"
    Set rsIngresos = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
    If Not (rsIngresos.EOF And rsIngresos.BOF) Then
      lIngresos = True
      lRsIngresos = True
    Else
      lIngresos = False
    End If
  End If
  
  'Selecciona Conceptos de Desctos
  If lDesctos Then
    sSQL = "SELECT res.codcpc, cpc.descpc, res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS importe "
    sSQL = sSQL & "FROM plresultado res "
    sSQL = sSQL & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    sSQL = sSQL & "AND res.codpdo='" & sPeriodo & "' "
    sSQL = sSQL & "AND res.codpsn='" & sPersona & "' "
    sSQL = sSQL & "AND res.tipocpc='" & s_Estado_Act & "' "
    sSQL = sSQL & "AND res.impbolecpc='" & s_Estado_Act & "' "
    sSQL = sSQL & "ORDER BY res.secuencia"
    Set rsDesctos = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
    If Not (rsDesctos.EOF And rsDesctos.BOF) Then
      lDesctos = True
      lRsDesctos = False
    Else
      lDesctos = False
    End If
  End If
  
  'Selecciona Conceptos de Aportes
  If lAportes Then
  
    
    sSQL = "SELECT res.codcpc, cpc.descpc, res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS importe "
    sSQL = sSQL & "FROM plresultado res "
    sSQL = sSQL & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    sSQL = sSQL & "AND res.codpdo='" & sPeriodo & "' "
    sSQL = sSQL & "AND res.codpsn='" & sPersona & "' "
    sSQL = sSQL & "AND res.tipocpc='" & s_Estado_Blq & "' "
    sSQL = sSQL & "AND res.impbolecpc='" & s_Estado_Act & "' "
    sSQL = sSQL & "ORDER BY res.secuencia"
    
    
    Set rsAportes = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
    If Not (rsAportes.EOF And rsAportes.BOF) Then
      lAportes = True
      lRsAportes = False
    Else
      lAportes = False
    End If
  End If
  
  'Imprime Detalle
  lProcesa = True
  Do While lProcesa
    'Ingreos
    If lIngresos Then
      If Not rsIngresos.EOF Then
        If nFila = nIngresos_CodigoY Or nFila = nIngresos_DescripcionY Or nFila = nIngresos_ImporteY Then
          'Imprime Linea
          If nIngresos_CodigoY <> 0 Then
            PrinterLine sIngresosC_FSize, sIngresosC_Font, sIngresosC_FontN, sIngresosC_FontS, sIngresosC_FontC, sIngresosC_FontTipo, rsIngresos("codcpc"), nIngresosC_Longitud, nIngresos_CodigoY, nIngresos_CodigoX, nLinea2ndFmt
          End If
          If nIngresos_DescripcionY <> 0 Then
            PrinterLine sIngresosD_FSize, sIngresosD_Font, sIngresosD_FontN, sIngresosD_FontS, sIngresosD_FontC, sIngresosD_FontTipo, rsIngresos("descpc"), nIngresosD_Longitud, nIngresos_DescripcionY, nIngresos_DescripcionX, nLinea2ndFmt
          End If
          If nIngresos_ImporteY <> 0 Then
            PrinterLine sIngresosI_FSize, sIngresosI_Font, sIngresosI_FontN, sIngresosI_FontS, sIngresosI_FontC, sIngresosI_FontTipo, rsIngresos("importe"), nIngresosI_Longitud, nIngresos_ImporteY, nIngresos_ImporteX, nLinea2ndFmt
          End If
          
          nIngresos_CodigoY = IIf(nIngresos_CodigoY <> 0, nIngresos_CodigoY + 1, 0)
          nIngresos_DescripcionY = IIf(nIngresos_DescripcionY <> 0, nIngresos_DescripcionY + 1, 0)
          nIngresos_ImporteY = IIf(nIngresos_ImporteY <> 0, nIngresos_ImporteY + 1, 0)
          rsIngresos.MoveNext
        End If
      Else
        lIngresos = False
      End If
    End If
    
    ' Descuentos
    If lDesctos Then
      If Not rsDesctos.EOF Then
        If nFila = nDesctos_CodigoY Or nFila = nDesctos_DescripcionY Or nFila = nDesctos_ImporteY Then
          'Imprime Linea
          If nDesctos_CodigoY <> 0 Then
            PrinterLine sDesctosC_FSize, sDesctosC_Font, sDesctosC_FontN, sDesctosC_FontS, sDesctosC_FontC, sDesctosC_FontTipo, rsDesctos("codcpc"), nDesctosC_Longitud, nDesctos_CodigoY, nDesctos_CodigoX, nLinea2ndFmt
          End If
          If nDesctos_DescripcionY <> 0 Then
            PrinterLine sDesctosD_FSize, sDesctosD_Font, sDesctosD_FontN, sDesctosD_FontS, sDesctosD_FontC, sDesctosD_FontTipo, rsDesctos("descpc"), nDesctosD_Longitud, nDesctos_DescripcionY, nDesctos_DescripcionX, nLinea2ndFmt
          End If
          If nDesctos_ImporteY <> 0 Then
            PrinterLine sDesctosI_FSize, sDesctosI_Font, sDesctosI_FontN, sDesctosI_FontS, sDesctosI_FontC, sDesctosI_FontTipo, rsDesctos("importe"), nDesctosI_Longitud, nDesctos_ImporteY, nDesctos_ImporteX, nLinea2ndFmt
          End If
          
          nDesctos_CodigoY = IIf(nDesctos_CodigoY <> 0, nDesctos_CodigoY + 1, 0)
          nDesctos_DescripcionY = IIf(nDesctos_DescripcionY <> 0, nDesctos_DescripcionY + 1, 0)
          nDesctos_ImporteY = IIf(nDesctos_ImporteY <> 0, nDesctos_ImporteY + 1, 0)
          rsDesctos.MoveNext
        End If
      Else
        lDesctos = False
      End If
    End If
  
    'Aportes
    If lAportes Then
      If Not rsAportes.EOF Then
        If nFila = nAportes_CodigoY Or nFila = nAportes_DescripcionY Or nFila = nAportes_ImporteY Then
          'Imprime Linea
          If nAportes_CodigoY <> 0 Then
            PrinterLine sAportesC_FSize, sAportesC_Font, sAportesC_FontN, sAportesC_FontS, sAportesC_FontC, sAportesC_FontTipo, rsAportes("codcpc"), nAportesC_Longitud, nAportes_CodigoY, nAportes_CodigoX, nLinea2ndFmt
          End If
          If nAportes_DescripcionY <> 0 Then
            PrinterLine sAportesD_FSize, sAportesD_Font, sAportesD_FontN, sAportesD_FontS, sAportesD_FontC, sAportesD_FontTipo, rsAportes("descpc"), nAportesD_Longitud, nAportes_DescripcionY, nAportes_DescripcionX, nLinea2ndFmt
          End If
          If nAportes_ImporteY <> 0 Then
            PrinterLine sAportesI_FSize, sAportesI_Font, sAportesI_FontN, sAportesI_FontS, sAportesI_FontC, sAportesI_FontTipo, rsAportes("importe"), nAportesI_Longitud, nAportes_ImporteY, nAportes_ImporteX, nLinea2ndFmt
          End If
          
          nAportes_CodigoY = IIf(nAportes_CodigoY <> 0, nAportes_CodigoY + 1, 0)
          nAportes_DescripcionY = IIf(nAportes_DescripcionY <> 0, nAportes_DescripcionY + 1, 0)
          nAportes_ImporteY = IIf(nAportes_ImporteY <> 0, nAportes_ImporteY + 1, 0)
          rsAportes.MoveNext
        End If
      Else
        lAportes = False
      End If
    End If
  
    If Not (lIngresos Or lDesctos Or lAportes) Then
      lProcesa = False
    End If
    nFila = nFila + 1
  Loop
  'Cerrar Cursores
  If lRsIngresos Then rsIngresos.Close
  If lRsDesctos Then rsDesctos.Close
  If lRsAportes Then rsAportes.Close

End Sub
Private Sub ImprimeFormato(sFormato As String, sPersona As String, sPeriodo As String, nCopias As Integer)
  Dim rs As New ADODB.Recordset
  Dim rsFormato As New ADODB.Recordset
  Dim rsPeriodo As New ADODB.Recordset
  Dim rsBoleta As New ADODB.Recordset
  
  Dim sSQL As String, sDato As String
  Dim dDato As String, sAtributo As String
  Dim nDato As Double
  Dim vAtributo As Variant
  
  Dim nPapelAlto As Double, nPapelAncho As Double
  Dim sFont As String, sCopia As String
  Dim nInicioCopia As Integer, nOrientacion As Integer
  Dim nCalidad As Integer
  
  Dim nTotalValores As Integer
  Dim aValores(), aTitulos()
  
  Dim nPos As Integer, nLinea2ndFmt As Integer, nFmt As Integer

  'Inicializa Variables
  nLinea2ndFmt = 0

  'Formato de Boleta
  sSQL = "SELECT * FROM plboletapago "
  sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND codboleta = '" & sFormato & "' "
  Set rsBoleta = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (rsBoleta.EOF And rsBoleta.BOF) Then
    nPapelAlto = CDec(rsBoleta("papelalto"))
    nPapelAncho = CDec(rsBoleta("papelancho"))
    sFont = rsBoleta("font")
    sCopia = rsBoleta("copia")
    nInicioCopia = CInt(rsBoleta("lininicopia"))
    nOrientacion = rsBoleta("orientacion")
    nCalidad = rsBoleta("calidad")
    rsBoleta.Close
  End If
  
  'Características de Datos Numericos
  nTotalValores = 0
  sSQL = "SELECT COUNT(dato) FROM pldetaboleta fbl "
  sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND fbl.codboleta='" & sFormato & "' "
  sSQL = sSQL & "AND fbl.origen = 'C'"
  Set rsFormato = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (rsFormato.EOF And rsFormato.BOF) Then
    nTotalValores = rsFormato(0)
    rsFormato.Close
  End If
  ReDim aValores(1 To nTotalValores)
  ReDim aTitulos(1 To nTotalValores)
  
  sSQL = "SELECT dato FROM pldetaboleta fbl "
  sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND fbl.codboleta='" & sFormato & "' "
  sSQL = sSQL & "AND fbl.origen = 'C'"
  Set rsFormato = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (rsFormato.EOF And rsFormato.BOF) Then
     nPos = 1
    Do While Not rsFormato.EOF
      ' Inicializo la información de conceptos
      aTitulos(nPos) = "C" & rsFormato("dato")
      aValores(nPos) = CDec(0)
      ' Obtengo el importe del concepto
      sSQL = "SELECT IFNULL(importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS importe "
      sSQL = sSQL & "FROM plresultado "
      sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
      sSQL = sSQL & "AND codpdo='" & sPeriodo & "' "
      sSQL = sSQL & "AND codpsn='" & sPersona & "' "
      sSQL = sSQL & "AND codcpc='" & rsFormato("dato") & "'"
      Set rs = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
      If Not (rs.EOF And rs.BOF) Then
        aTitulos(nPos) = "C" & rsFormato("dato")
        aValores(nPos) = CDec(rs("importe"))
        rs.Close
      End If
      nPos = nPos + 1
      rsFormato.MoveNext
    Loop
    rsFormato.Close
  End If
  
  'Datos generales del personal
  sSQL = "SELECT DISTINCTROW "
  sSQL = sSQL & "res.codpsn AS Codigo_Personal, "
  sSQL = sSQL & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS Apellido_Nombre, "
  sSQL = sSQL & "dxr.fecingreso AS Fecha_Ingreso, "
  sSQL = sSQL & "IFNULL(dxr.naciextrapsn, '0') AS Domiciliado, "
  sSQL = sSQL & "IFNULL(dxr.codcgo, '') AS Codigo_Cargo, "
  sSQL = sSQL & "IFNULL(dxr.codcco, '') AS Codigo_CenCosto, "
  sSQL = sSQL & "psn.pagodolar AS Pago_Dolares, "
  sSQL = sSQL & "IFNULL(dxr.estadopsn, 'A') AS Estado, "
  sSQL = sSQL & "IFNULL(dxr.codafp, '') AS Codigo_AFP, "
  sSQL = sSQL & "IFNULL(psn.numeroafp, '') AS Numero_AFP, "
  sSQL = sSQL & "psn.fecnacimiento AS Fecha_Nacimiento, "
  sSQL = sSQL & "CONCAT(IFNULL(psn.nomviadirec, ''), ' ', IFNULL(psn.numerdirec, ''), ' ', IFNULL(psn.nomzondirec, '')) AS Direccion, "
  sSQL = sSQL & "IFNULL(psn.telefono, '') AS Telefono, "
  sSQL = sSQL & "psn.estcivilpsn AS Estado_Civil, "
  sSQL = sSQL & "(CASE psn.sexopsn WHEN '0' THEN 'M' ELSE 'F' END) AS Sexo, "
  sSQL = sSQL & "psn.numdepen AS Carga_Familiar, "
  sSQL = sSQL & "IFNULL(dxr.codeps, '') AS Codigo_EPS, "
  sSQL = sSQL & "IFNULL(psn.numdociden, '') AS Doc_Identidad, "
  sSQL = sSQL & "IFNULL(psn.numdocmil, '') AS Doc_Militar, "
  sSQL = sSQL & "IFNULL(psn.codbcopago, '') AS Banco_Pago, "
  sSQL = sSQL & "IFNULL(psn.cuentapago, '') AS Cuenta_pago, "
  sSQL = sSQL & "IFNULL(psn.jornadalaboral, '') AS Jornada_laboral, "
  sSQL = sSQL & "res.codpdo AS Cod_Periodo, "
  sSQL = sSQL & "pdo.despdo AS Des_Periodo, "
  sSQL = sSQL & "IFNULL(psn.nroessalud, '') AS Nro_Essalud, "
  sSQL = sSQL & "IFNULL(asi.diatrabajo, 0) AS Dias_Trabajado, "
  sSQL = sSQL & "IFNULL(asi.diamediotm, 0) AS Dias_MedioTiempo, "
  sSQL = sSQL & "IFNULL(asi.diaparcial, 0) AS Dias_TiempoParcial, "
  sSQL = sSQL & "IFNULL(asi.dialaboral, 0) AS Dias_Laborado, "
  sSQL = sSQL & "IFNULL(asi.diaferiado, 0) AS Dias_Feriado, "
  sSQL = sSQL & "IFNULL(asi.horanormal, 0) AS Hora_Normal, "
  sSQL = sSQL & "IFNULL(asi.horamediotm, 0) AS Hora_MedioTiempo, "
  sSQL = sSQL & "IFNULL(asi.horaparcial, 0) AS Hora_TiempoParcial, "
  sSQL = sSQL & "asi.fechainivacacion AS FecIni_Vacaciones, "
  sSQL = sSQL & "asi.fechafinvacacion AS FecFin_Vacaciones, "
  sSQL = sSQL & "asi.fechacese AS Fecha_Cese, "
  sSQL = sSQL & "IFNULL(cgo.descgo, '') AS Descrip_Cargo, "
  sSQL = sSQL & "IFNULL(afp.desafp, '') AS Descrip_AFP, "
  sSQL = sSQL & "IFNULL(eps.deseps, '') AS Descrip_EPS, "
  sSQL = sSQL & "IFNULL(cco.detcco, '') AS Descrip_CenCosto, "
  sSQL = sSQL & "IFNULL(psn.codtpt, '') AS Tipo_Personal, "
  sSQL = sSQL & "IFNULL(tpt.destpt, '') AS Descrip_TipoPsn, "
  sSQL = sSQL & "IFNULL(dxr.codubica, '') AS Cod_Ubicacion, "
  sSQL = sSQL & "IFNULL(dxr.codsec, '') AS Cod_Seccion, "
  sSQL = sSQL & "IFNULL(ubi.desubica, '') AS Descrip_Ubicacion, "
  sSQL = sSQL & "IFNULL(sec.dessec, '') AS Descrip_Seccion, "
  sSQL = sSQL & "IFNULL(asi.horatipo1, 0) AS Hora_ExtraSimple, "
  sSQL = sSQL & "IFNULL(asi.horatipo2, 0) AS Hora_ExtraDoble, "
  sSQL = sSQL & "IFNULL(asi.horatipo3, 0) AS Hora_Especial, "
  sSQL = sSQL & "IFNULL(asi.horatipo4, 0) AS Hora_Nocturna, "
  sSQL = sSQL & "pdo.fechaini AS FecIni_Periodo, "
  sSQL = sSQL & "pdo.fechafin AS FecFin_Periodo, "
  sSQL = sSQL & "pdo.tipocambio AS Tipo_Cambio, "
  sSQL = sSQL & "IFNULL(asi.diavacaciones, 0) AS Dias_Vacaciones, "
  sSQL = sSQL & "(CASE psn.cgoconfianza WHEN '1' THEN 'PERSONAL DE CONFIANZA' WHEN '2' THEN 'PERSONAL DE DIRECCION' WHEN '3' THEN 'PERSONAL NO S.F.I.' WHEN '4' THEN 'STAFF' WHEN '5' THEN 'PERSONAL DE CONFIANZA NO S.F.I.' ELSE '' END) AS Descrip_Confianza, "
  sSQL = sSQL & "REPEAT(' ', 5) AS Texto "
  For nPos = 1 To nTotalValores
    If aTitulos(nPos) <> "" Then
      sSQL = sSQL & ", "
      sSQL = sSQL & aValores(nPos) & " AS " & aTitulos(nPos) & " "
    End If
  Next
  sSQL = sSQL & "FROM plpersonal psn "
  sSQL = sSQL & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
  sSQL = sSQL & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSQL = sSQL & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  sSQL = sSQL & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSQL = sSQL & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  sSQL = sSQL & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  sSQL = sSQL & "LEFT JOIN pltpotrabajador tpt ON psn.codtpt=tpt.codtpt "
  sSQL = sSQL & "LEFT JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSQL = sSQL & "LEFT JOIN plentidadeps eps ON dxr.codeps=eps.codeps "
  sSQL = sSQL & "LEFT JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSQL = sSQL & "LEFT JOIN plseccion sec ON dxr.codsec=sec.codsec "
  sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  If sPersona <> "" Then
    sSQL = sSQL & "AND psn.codpsn= '" & sPersona & "' "
  End If
  sSQL = sSQL & "AND res.codpdo='" & sPeriodo & "'"
  
  ' Setea Caracerísticas de cada línea
  With Printer
    .ScaleMode = vbCentimeters
    .Orientation = nOrientacion
    .PrintQuality = nCalidad
    .Width = CDec(nPapelAncho * 567)
    .Height = CDec(nPapelAlto * 567)
    .Font = sFont
    .Copies = nCopias
  End With
  Set rs = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  
  If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
      For nFmt = 0 To CInt(sCopia)
        If nFmt = 0 Then
          nLinea2ndFmt = 0
        Else
          nLinea2ndFmt = nInicioCopia
        End If
        'Datos de Cabecera
        sSQL = "SELECT seccion, dato, IFNULL(desdato, '') AS desdato, tipodato, nombre, "
        sSQL = sSQL & "fila, columna, longitud, sizefont, fontn, fonts, fontc "
        sSQL = sSQL & "FROM pldetaboleta fbl "
        sSQL = sSQL & "LEFT JOIN plvarfunc pvf ON pvf.codigo = fbl.dato "
        sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' AND codboleta='" & sFormato & "' AND seccion='C' "
        sSQL = sSQL & "ORDER BY fila, columna"
        Set rsFormato = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
        If Not (rsFormato.EOF And rsFormato.BOF) Then
          Do While Not rsFormato.EOF
            'Captura Dato
            sAtributo = IIf(IsNull(rsFormato("nombre")), "C" & rsFormato("dato"), rsFormato("nombre"))
            vAtributo = IIf(rsFormato("desdato") = "", rs(sAtributo), rsFormato("desdato"))
            'Imprime Línea
            PrinterLine rsFormato("sizefont"), sFont, rsFormato("fontn"), rsFormato("fonts"), rsFormato("fontc"), rsFormato("tipodato"), vAtributo, rsFormato("longitud"), rsFormato("fila"), rsFormato("columna"), nLinea2ndFmt
            rsFormato.MoveNext
          Loop
          rsFormato.Close
        End If
        
        'Datos de Detalle
        Imprime_Detalle sFormato, sPersona, sPeriodo, sFont, nLinea2ndFmt
        
        'Datos de Pie
        sSQL = "SELECT seccion, dato,IFNULL(desdato, '') AS desdato, tipodato, nombre, fila, columna, longitud, sizefont, fontn, fonts, fontc "
        sSQL = sSQL & "FROM pldetaboleta fbl "
        sSQL = sSQL & "LEFT JOIN plvarfunc pvf ON pvf.codigo = fbl.dato "
        sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' AND codboleta='" & sFormato & "' AND seccion='P' "
        sSQL = sSQL & "ORDER BY fila, columna"
        Set rsFormato = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
        If Not (rsFormato.EOF And rsFormato.BOF) Then
          Do While Not rsFormato.EOF
            'Captura Dato
            sAtributo = IIf(IsNull(rsFormato("nombre")), "C" & rsFormato("dato"), rsFormato("nombre"))
            vAtributo = IIf(rsFormato("desdato") = "", rs(sAtributo), rsFormato("desdato"))
            'Imprime Línea
            PrinterLine rsFormato("sizefont"), sFont, rsFormato("fontn"), rsFormato("fonts"), rsFormato("fontc"), rsFormato("tipodato"), vAtributo, rsFormato("longitud"), rsFormato("fila"), rsFormato("columna"), nLinea2ndFmt
            rsFormato.MoveNext
          Loop
          rsFormato.Close
        End If
      Next
      'Lee Siguiente Persona
      rs.MoveNext
      'Nueva Hoja
      Printer.NewPage
    Loop
    rs.Close
  End If

End Sub
Private Sub PrinterLine(sSizeFont As String, sFont As String, sFontN As String, sFontS As String, sFontC As String, sTipoDato As String, vDato As Variant, nLongitud As Integer, nFila As Integer, nColumna As Integer, nLinea2ndFmt As Integer)
  Dim sDatoPrint As String
  Dim nDatoPrint As Double

  With Printer
  
    'Personalizar Impresión de Línea
    .ScaleMode = vbCharacters
    .FontSize = CDec(sSizeFont)
    .Font.Bold = (sFontN = s_Estado_Act)
    .Font.Underline = (sFontS = s_Estado_Act)
    .Font.Italic = (sFontC = s_Estado_Act)
  
    'Impresión de Caracteres
    If sTipoDato = "C" Then
      sDatoPrint = Mid(IIf(IsNull(vDato), "", vDato), 1, nLongitud)
      .CurrentY = nFila + nLinea2ndFmt: .CurrentX = nColumna
      Printer.Print sDatoPrint
    End If
    
    'Impresión de Números
    If sTipoDato = "N" Then
      nDatoPrint = CDec(vDato)
      .CurrentY = nFila + nLinea2ndFmt: .CurrentX = nColumna
      sDatoPrint = gdl_Funcion.PadL(FormatNumber(nDatoPrint, 2), nLongitud, " ")
      Printer.Print sDatoPrint
    End If
    
    'Impresión de Fechas
    If sTipoDato = "F" Then
      sDatoPrint = IIf(IsNull(vDato), "", vDato)
      .CurrentY = CInt(nFila) + nLinea2ndFmt: .CurrentX = CInt(nColumna)
      Printer.Print Format(sDatoPrint, s_FormatoFecha)
    End If
  End With

End Sub
Private Sub PrnBoletaGenerica(ByVal nTabIndex As Integer, ByVal s_Proceso As String, ByVal s_FechaHora As String, ByVal s_Tabla As String, ByVal n_Accion As Integer)
  Dim porstDetalle As ADODB.Recordset
  Dim a_Detalle(), nContador As Integer
  Dim nColumna As Integer, sColumna As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim a_Boleta(9), nDetalle As Integer, nCopia As Integer
  Dim s_DesConfianza As String
  Dim sMoneda As String, sMonedaPago As String
  Dim nImporteIng As Double, nImporteDsc As Double, nImportePag As Double
  Dim nSalIniCtaCte As Double, nSalFinCtaCte As Double, nDsctoCtaCte As Double
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Imprimiendo Boletas ..."

  '[ Genero la tabla temporal de datosdel personal
  s_Sql = "DROP TABLE IF EXISTS tmpdatospsn"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  ' Genero la informacion de datos personales
  s_Sql = "CREATE TEMPORARY TABLE tmpdatospsn "
  s_Sql = s_Sql & "SELECT DISTINCT psn.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, psn.cgoconfianza, "
  s_Sql = s_Sql & "IFNULL(dxr.codcgo, '') AS codcgo, IFNULL(cgo.descgo, '') AS descgo, dxr.fecingreso, IFNULL(psn.nroessalud, '') AS nroessalud, "
  s_Sql = s_Sql & "IFNULL(psn.numdociden, '') AS numdociden, IFNULL(psn.numeroafp, '') AS numeroafp, IFNULL(dxr.codafp, '') AS codafp, IFNULL(afp.desafp, '') AS desafp, "
  s_Sql = s_Sql & "(asi.diatrabajo+asi.diamediotm+asi.diaparcial+asi.licencia) AS diatrabajo, asi.dialaboral, asi.diaferiado, asi.diafalta, "
  s_Sql = s_Sql & "(asi.diaprepostnatal+asi.accidente) AS diasubsidio, (asi.licencia+asi.diavacaciones) AS dianolaboral, asi.enfermedad, "
  s_Sql = s_Sql & "asi.diavacaciones, asi.diatradesemanal, asi.diasuspension, asi.dialibre, asi.licencia, "
  s_Sql = s_Sql & "(asi.horatipo1 + asi.horatipo2 + asi.horatipo3 + asi.horatipo4) AS horaextra, asi.horatipo1, asi.horatipo2, asi.horatipo3, "
  s_Sql = s_Sql & "asi.horatipo4, (asi.opcional*0) AS opcional, (asi.horanormal+asi.horamediotm+asi.horaparcial) AS horanormal, asi.fechainivacacion, asi.fechafinvacacion, asi.fechacese, "
  s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', '" & s_Codmon_me & "', '" & s_Codmon_mn & "') AS monpago , IFNULL(pdo.despdo, '') AS despdo, pdo.tipocambio, "
  s_Sql = s_Sql & "000000000000000000.00 AS basico, "
  s_Sql = s_Sql & "cco.detcco, "
  s_Sql = s_Sql & "pdo.fechaini, pdo.fechafin, "
  s_Sql = s_Sql & "IFNULL(cdt.descdt, '') AS descdt, "
  s_Sql = s_Sql & "doc.desdci as desdci, tra.destpt as destpt, "
  s_Sql = s_Sql & "asi.tardanza as tardanza, sec.dessec as seccion, ubi.desubica AS ubicacion, "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "res.codpsn", "dxr.codcco ", "dxr.codubica", "dxr.codsec ") & " AS codtab "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad doc ON psn.coddci=doc.coddci "
  s_Sql = s_Sql & "LEFT JOIN pltpotrabajador tra ON psn.codtpt=tra.codtpt "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "LEFT JOIN plconditrabajo cdt ON dxr.codcls=cdt.codcls AND dxr.codcdt=cdt.codcdt "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
  s_Sql = s_Sql & "AND dxr." & Choose(nTabIndex + 1, "codpsn", "codcco", "codubica", "codsec") & " IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  If n_Accion = 10 Then
    s_Sql = s_Sql & "AND IFNULL(psn.correoelect, '')<>'' "
  End If
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "", "dxr.codcco, ", "dxr.codubica, ", "dxr.codsec, ") & "codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Actualizo el importe del basico
  s_Sql = "UPDATE tmpdatospsn tmp, plresultado res, plparametroafp par "
  s_Sql = s_Sql & "SET tmp.basico=res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
  s_Sql = s_Sql & "AND res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & "AND res.pdoano=par.pdoano "
  s_Sql = s_Sql & "AND res.codcpc=par.cpcbasico"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar

  ' Obtengo los datos del personal
  s_Sql = "SELECT psn.codpsn, psn.nompsn, psn.cgoconfianza, psn.codcgo, psn.descgo, psn.fecingreso, psn.nroessalud, psn.numdociden, "
  s_Sql = s_Sql & "psn.numeroafp, psn.codafp, psn.desafp, psn.diatrabajo, psn.dialaboral, psn.diaferiado, psn.diafalta, psn.diasubsidio, "
  s_Sql = s_Sql & "psn.dianolaboral, psn.enfermedad, psn.diavacaciones, psn.diatradesemanal, psn.diasuspension, psn.dialibre, psn.licencia, "
  s_Sql = s_Sql & "psn.horaextra, psn.horatipo1, psn.horatipo2, psn.horatipo3, "
  s_Sql = s_Sql & "psn.horatipo4, psn.opcional, psn.horanormal, psn.fechainivacacion, psn.fechafinvacacion, psn.fechacese, psn.monpago, psn.despdo, psn.tipocambio, "
  s_Sql = s_Sql & "psn.basico, psn.detcco, psn.fechaini, psn.fechafin, "
  s_Sql = s_Sql & "psn.descdt, psn.desdci, psn.destpt, "
  s_Sql = s_Sql & "psn.tardanza, psn.seccion, psn.ubicacion, psn.codtab "
  s_Sql = s_Sql & "FROM tmpdatospsn psn "
  s_Sql = s_Sql & "ORDER BY psn.codtab, psn.codpsn"
  

  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    fMenu.panPercent.Visible = True
    sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    Do While Not porstRecordset.EOF
      ' Obtengo saldo de cuenta corriente
      nDsctoCtaCte = 0
      nSalIniCtaCte = gdl_Funcion.DameCuentaCorriente(gdl_Conexion.CadenaConexion, ps_ClsPlanilla, porstRecordset("codpsn"), txtPeriodo.Text, porstRecordset!monpago, nDsctoCtaCte)
      nSalFinCtaCte = Round(nSalIniCtaCte - nDsctoCtaCte, 2)
      s_DesConfianza = Choose(CInt(porstRecordset!cgoconfianza) + 1, "", "PERSONAL DE CONFIANZA", "PERSONAL DE DIRECCION", "PERSONAL NO S.F.I.", "STAFF", "PERSONAL DE CONFIANZA NO S.F.I.")
      sMonedaPago = IIf(porstRecordset!monpago = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt)
      ' Obtengo el detalle del Cálculo
      s_Sql = "SELECT res.codpsn, res.secuencia, res.codcpc, cpc.descpc, res.tipocpc, "
      s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Ina & "', res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS imporingreso, "
      s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Act & "', res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS impordescto, "
      s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Blq & "', res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS imporaporte, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & porstRecordset("codpsn") & "' "
      s_Sql = s_Sql & "ORDER BY tipocpc, secuencia"
      Set porstDetalle = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      nImporteIng = 0: nImporteDsc = 0: nImportePag = 0
      n_Index = 0: nContador = 0: nColumna = 9
      ReDim a_Detalle(9, 0)
      Do While Not porstDetalle.EOF
        ' selecciono el tipo de concepto
        If nColumna <> CInt(porstDetalle("tipocpc")) Then
          nColumna = CInt(porstDetalle("tipocpc"))
          sColumna = Choose(nColumna + 1, "imporingreso", "impordescto", "imporaporte")
          nContador = 0
        End If
        nContador = nContador + 1
        ' Redimensiono e inicializo el arreglo de los detalles
        If nContador > UBound(a_Detalle, 2) Then
          ReDim Preserve a_Detalle(9, nContador)
          a_Detalle(1, nContador) = "": a_Detalle(2, nContador) = "": a_Detalle(3, nContador) = ""
          a_Detalle(4, nContador) = "": a_Detalle(5, nContador) = "": a_Detalle(6, nContador) = ""
          a_Detalle(7, nContador) = CDec(0): a_Detalle(8, nContador) = CDec(0): a_Detalle(9, nContador) = CDec(0)
        End If
        ' Asigno los datos al arreglo
        a_Detalle(nColumna + 1, nContador) = porstDetalle("codcpc")
        a_Detalle(nColumna + 4, nContador) = porstDetalle("descpc")
        a_Detalle(nColumna + 7, nContador) = CDec(porstDetalle(sColumna))
        ' Obtengo ingresos y descuentos otra moneda
        nImporteIng = nImporteIng + CDec(Choose(nColumna + 1, porstDetalle!importecmb, 0, 0))
        nImporteDsc = nImporteDsc + CDec(Choose(nColumna + 1, 0, porstDetalle!importecmb, 0))
        porstDetalle.MoveNext
      Loop
      porstDetalle.Close
      ' Obtengo el importe en otra moneda
      nImportePag = Round(nImporteIng - nImporteDsc, 2)
      nImportePag = IIf(sMoneda = sMonedaPago, 0, nImportePag)
      
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      ' Inserto los detalle maximo 10
      nDetalle = IIf(UBound(a_Detalle, 2) > 10, UBound(a_Detalle, 2), 10)
      For n_Index = 1 To nDetalle
        ' Inicializo los datos del detalle
        a_Boleta(1) = "": a_Boleta(2) = "": a_Boleta(3) = ""
        a_Boleta(4) = "": a_Boleta(5) = "": a_Boleta(6) = ""
        a_Boleta(7) = 0: a_Boleta(8) = 0: a_Boleta(9) = 0
        If UBound(a_Detalle, 2) >= n_Index Then
          a_Boleta(1) = a_Detalle(1, n_Index): a_Boleta(2) = a_Detalle(2, n_Index)
          a_Boleta(3) = a_Detalle(3, n_Index): a_Boleta(4) = a_Detalle(4, n_Index)
          a_Boleta(5) = a_Detalle(5, n_Index): a_Boleta(6) = a_Detalle(6, n_Index)
          a_Boleta(7) = a_Detalle(7, n_Index): a_Boleta(8) = a_Detalle(8, n_Index)
          a_Boleta(9) = a_Detalle(9, n_Index)
        End If
        a_Campos = Array("codpsn", "nompsn", "codcgo", "descgo", "fecingreso", "nroessalud", "despdo", "desconfianza", _
                        "fechainivacacion", "fechafinvacacion", "numdociden", "numeroafp", "codafp", "desafp", "diatrabajo", "diasubsidio", "diaferiado", "diafalta", "dialaboral", "dianolaboral", _
                        "horanormal", "horaextra", "horaextra1", "horaextra2", "horaextra3", "horaextra4", "horaextrax", "fecbaja", "copia", "secuencia", "basico", "codcpcing", "descpcing", "impcpcing", "codcpcdsc", _
                        "descpcdsc", "impcpcdsc", "codcpcapo", "descpcapo", "impcpcapo", "moneda", "monpago", "importipcmb", "impornetocmb", "salinicte", "salfincte", "detcco", "fechaini", "fechafin", _
                        "descdt", "tipdoc", "tiptra", "tardanza", "diavacaciones", "diatradesemanal", "diasuspension", "dialibre", "dialicencia", "diaenfermedad", "seccion", "ubicacion", "codtab")
        a_Valores = Array(porstRecordset("codpsn"), porstRecordset("nompsn"), porstRecordset("codcgo"), porstRecordset("descgo"), Format(porstRecordset("fecingreso"), s_FmtFechMysql_0), porstRecordset("nroessalud"), porstRecordset("despdo"), s_DesConfianza, _
                          Format(porstRecordset("fechainivacacion"), s_FmtFechMysql_0), Format(porstRecordset("fechafinvacacion"), s_FmtFechMysql_0), porstRecordset("numdociden"), porstRecordset("numeroafp"), porstRecordset("codafp"), porstRecordset("desafp"), CDec(porstRecordset("diatrabajo")), CDec(porstRecordset("diasubsidio")), CDec(porstRecordset("diaferiado")), CDec(porstRecordset("diafalta")), CDec(porstRecordset("dialaboral")), CDec(porstRecordset("dianolaboral")), _
                          CDec(porstRecordset("horanormal")), CDec(porstRecordset("horaextra")), CDec(porstRecordset("horatipo1")), CDec(porstRecordset("horatipo2")), CDec(porstRecordset("horatipo3")), CDec(porstRecordset("horatipo4")), CDec(porstRecordset("opcional")), Format(porstRecordset("fechacese"), s_FmtFechMysql_0), s_Estado_Ina, n_Index, CDec(porstRecordset("basico")), a_Boleta(1), a_Boleta(4), a_Boleta(7), a_Boleta(2), _
                          a_Boleta(5), a_Boleta(8), a_Boleta(3), a_Boleta(6), a_Boleta(9), sMoneda, sMonedaPago, CDec(porstRecordset("tipocambio")), nImportePag, nSalIniCtaCte, nSalFinCtaCte, porstRecordset("detcco"), Format(porstRecordset("fechaini"), s_FmtFechMysql_0), Format(porstRecordset("fechafin"), s_FmtFechMysql_0), _
                          porstRecordset("descdt"), porstRecordset("desdci"), porstRecordset("destpt"), CDec(porstRecordset("tardanza")), CDec(porstRecordset("diavacaciones")), CDec(porstRecordset("diatradesemanal")), CDec(porstRecordset("diasuspension")), CDec(porstRecordset("dialibre")), CDec(porstRecordset("licencia")), CDec(porstRecordset("enfermedad")), porstRecordset("seccion"), porstRecordset("ubicacion"), porstRecordset!codtab)
        a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
                        TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, _
                        TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, _
                        TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, _
                        TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
        ' Realizo la actualización de los registros
        For nCopia = 0 To IIf(ribCopia.Value, 1, 0)
          a_Valores(28) = Choose(nCopia + 1, s_Estado_Ina, s_Estado_Act)
          If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        Next nCopia
      Next n_Index
      gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Loop
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Elimino la tabla temporal
  gdl_Conexion.Execucion "DROP TABLE IF EXISTS tmpdatospsn", Elimina
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub PrnBoletaFormato(nTabIndex As Integer, ByVal s_Proceso As String, ByVal s_FechaHora As String, ByVal n_Copias As Integer)
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Imprimiendo Boletas ..."
  
  '[ Genero la tabla temporal de datosdel personal
  s_Sql = "DROP TABLE IF EXISTS tmpdatospsn"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  ' Obtengo los datos del personal
  s_Sql = "CREATE TEMPORARY TABLE tmpdatospsn "
  s_Sql = s_Sql & "SELECT DISTINCT psn.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, psn.cgoconfianza, "
  s_Sql = s_Sql & "IFNULL(dxr.codcgo, '') AS codcgo, IFNULL(cgo.descgo, '') AS descgo, dxr.fecingreso, IFNULL(psn.nroessalud, '') AS nroessalud, "
  s_Sql = s_Sql & "IFNULL(psn.numdociden, '') AS numdociden, IFNULL(psn.numeroafp, '') AS numeroafp, IFNULL(dxr.codafp, '') AS codafp, "
  s_Sql = s_Sql & "IFNULL(afp.desafp, '') AS desafp, asi.diatrabajo, asi.horanormal, asi.fechainivacacion, asi.fechafinvacacion, asi.fechacese, "
  s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', '" & s_Codmon_me & "', '" & s_Codmon_mn & "') AS monpago , IFNULL(pdo.despdo, '') AS despdo, pdo.tipocambio, "
  s_Sql = s_Sql & "000000000000000000.00 AS basico "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "", ", dxr.codcco ", ", dxr.codubica ", ", dxr.codsec ")
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
  s_Sql = s_Sql & "AND dxr." & Choose(nTabIndex + 1, "codpsn", "codcco", "codubica", "codsec") & " IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "", "dxr.codcco, ", "dxr.codubica, ", "dxr.codsec, ") & IIf(ribOrdenar.Value, "nompsn", "codpsn")
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Actualizo el importe del basico
  s_Sql = "UPDATE tmpdatospsn tmp, plresultado res, plparametroafp par "
  s_Sql = s_Sql & "SET tmp.basico=res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
  s_Sql = s_Sql & "AND res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & "AND res.pdoano=par.pdoano "
  s_Sql = s_Sql & "AND res.codcpc=par.cpcbasico"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar

  ' Obtengo los datos del personal
  s_Sql = "SELECT psn.codpsn, psn.nompsn, psn.cgoconfianza, psn.codcgo, psn.descgo, psn.fecingreso, "
  s_Sql = s_Sql & "psn.nroessalud, psn.numdociden, psn.numeroafp, psn.codafp, psn.desafp, psn.diatrabajo, "
  s_Sql = s_Sql & "psn.horanormal, psn.fechainivacacion, psn.fechafinvacacion, psn.fechacese, psn.monpago , "
  s_Sql = s_Sql & "psn.despdo, psn.tipocambio, psn.basico "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "", ", psn.codcco ", ", psn.codubica ", ", psn.codsec ")
  s_Sql = s_Sql & "FROM tmpdatospsn psn "
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "", "psn.codcco, ", "psn.codubica, ", "psn.codsec, ") & IIf(ribOrdenar.Value, "psn.nompsn", "psn.codpsn")
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    Do While Not porstRecordset.EOF
      ' Imprime la boleta de pago
      For n_Index = 1 To n_Copias
        ImprimeFormato txtFormato.Text, porstRecordset("codpsn"), txtPeriodo.Text, n_Copias
      Next n_Index
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Loop
  End If
  GoTo Finalizar
  
Finalizar:
  ' Elimino la tabla temporal
  gdl_Conexion.Execucion "DROP TABLE IF EXISTS tmpdatospsn", Elimina
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub RecuperaRegistros(ByVal nIndex As Integer, ByVal s_Orden As String)


  ' Recuperaron Información
  If nIndex = 0 Then
    s_Sql = "SELECT codcls, codpsn, apepaterno, apematerno, nombres, estadopsn "
    s_Sql = s_Sql & "FROM plpersonal "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    If Not ribParametro(0).Value Then
      s_Sql = s_Sql & "AND estadopsn" & IIf(ribParametro(1).Value, "<>", "=") & "'I' "
    End If
    s_Sql = s_Sql & " ORDER BY " & s_Orden
  ElseIf nIndex = 1 Then
    s_Sql = "SELECT codcco, detcco, estcco "
    s_Sql = s_Sql & "FROM cocco "
    s_Sql = s_Sql & "WHERE LENGTH(codcco)>=" & pn_NivelCenCosto & " "
    s_Sql = s_Sql & " ORDER BY " & s_Orden
  ElseIf nIndex = 2 Then
    s_Sql = "SELECT codubica, desubica, estadoubica "
    s_Sql = s_Sql & "FROM plubicacion "
    s_Sql = s_Sql & " ORDER BY " & s_Orden
  ElseIf nIndex = 3 Then
    s_Sql = "SELECT codsec, dessec, estadosec "
    s_Sql = s_Sql & "FROM plseccion "
    s_Sql = s_Sql & " ORDER BY " & s_Orden
  End If
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & IIf(nIndex = 1, ps_DaBasCon, ps_DataBase), dcaSeleccion(nIndex), tdbSeleccion(nIndex), s_Sql, adCmdText, adLockReadOnly
  ' Inicializo los rangos de impresion
  as_SelRegistro(nIndex, 0) = "": as_SelRegistro(nIndex, 1) = ""
  If dcaSeleccion(nIndex).Recordset.RecordCount > 0 Then
    dcaSeleccion(nIndex).Recordset.MoveLast: as_SelRegistro(nIndex, 1) = dcaSeleccion(nIndex).Recordset.Bookmark
    dcaSeleccion(nIndex).Recordset.MoveFirst: as_SelRegistro(nIndex, 0) = dcaSeleccion(nIndex).Recordset.Bookmark
  End If

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_Proceso As String, s_OldMessage As String
  Dim s_Expresion As String, sOrden As String, s_PdfWhere As String, s_TmpCarpeta As String
  Dim s_RegPatronal As String, s_Representante As String, sDireccion As String, s_Cargo As String
  Dim s_UsuarioEnvio As String, s_PasswordEnvio As String, s_CorreoEnvio As String
  Dim porstExporta As ADODB.Recordset
  Dim nTabIndex As Integer, n_ServerEnvio As Integer, n_PuertoEnvio As Integer
  Dim oCreadorPdf As Object
  Dim nTemporizador As Long
  
  nTabIndex = tabRegister.Tab
  ' Verifico que Existan Registros
  If (dcaSeleccion(0).Recordset.EOF Or dcaSeleccion(0).Recordset.BOF) Or (dcaSeleccion(0).Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  
  'Enero 2015
  If Valida_LicenciaUso(ps_Ano, ps_Fecha_LimiteProc, Right(Left(txtPeriodo, 6), 2), Left(txtPeriodo.Text, 4)) = False Then
    MsgBox "Operación no puede ser procesada" & Chr(13) & "Restricción de Licencia no permite este proceso" & Chr(13) & "Por favor comuniquese con el personal de Sistemas.", vbInformation
    Exit Sub
  End If
  
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
   Case 6, 7, 8, 9, 10 ' Opciones de impresión
    ' Verifico que existan registros seleccionados
    If txtFormato.Text = "" And Index = 8 Then Beep: MsgBox "Debe Ingresar el Formato de la boleta", vbExclamation: txtFormato.SetFocus: Exit Sub
    If (lblHelp(0).Caption = "" Or lblHelp(0).Caption = "???") And Index = 8 Then Beep: MsgBox "Formato de la boleta no existe; verifique", vbExclamation: txtFormato.SetFocus: Exit Sub
    If txtPeriodo.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(1).Caption = "" Or lblHelp(1).Caption = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    ' Verifico que existan registros seleccionados
    nTabIndex = tabRegister.Tab
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 1 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 2 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 3 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    
    ' Obtengo la direccion de la empresa
    sDireccion = ""
    s_Sql = "SELECT via.abrevia, cfg.direccionvia, cfg.numerodir, zon.abrezona, cfg.direccionzona, cfg.ubigeodir,cfg.firma,cfg.logo "
    s_Sql = s_Sql & "FROM plcfgempresa cfg "
    s_Sql = s_Sql & "LEFT JOIN pltipovia via ON cfg.codvia=via.codvia "
    s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON cfg.codzona=zon.codzona "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Ano & "' and cfg.dirimpbol=1 "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDireccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!direccionzona) & " - " & sDireccion
      
      
     'Enero 2015 Obtenfo el campo de la BD de logo y firma si esiste valor en dichos campos.
     'La imagen generada la guardo en un directorio en el disco duro.
    
        Dim m_stream    As ADODB.Stream
        'Obtengo Firma
        If IsNull(porstRecordset.Fields("firma").Value) = False Then
            Set m_stream = New ADODB.Stream
            m_stream.Type = adTypeBinary
            m_stream.Open
            m_stream.Write porstRecordset.Fields("firma").Value
            ' -- Guardar los datos en disco
            m_stream.SaveToFile "C:\tmpexportar\firma.bmp", adSaveCreateOverWrite
        End If
    
        'Obtengo Logo.
        If IsNull(porstRecordset.Fields("logo").Value) = False Then
            Set m_stream = New ADODB.Stream
            m_stream.Type = adTypeBinary
            m_stream.Open
            m_stream.Write porstRecordset.Fields("logo").Value
            ' -- Guardar los datos en disco
            m_stream.SaveToFile "C:\tmpexportar\logo.bmp", adSaveCreateOverWrite
        End If
    End If
    
   
  
    s_Expresion = IIf(ribFirma.Value, "ger", "rep")
    ' Verifico que existan parametros de boletas
    s_Sql = "SELECT DISTINCT cls.fmtboleta,  prm.cpcbasico, cfg.regpatronal, cfg.repimpbol, "
    s_Sql = s_Sql & "CONCAT(IFNULL(cfg." & s_Expresion & "apepaterno, ''), ' ', IFNULL(cfg." & s_Expresion & "apematerno, ''), ', ', "
    s_Sql = s_Sql & "IFNULL(cfg." & s_Expresion & "nombres, '')) AS representante, plcargo.descgo, "
    s_Sql = s_Sql & "IFNULL(cfg.server_envio, 0) AS server_envio, IFNULL(cfg.usuario_envio, '') AS usuario_envio, "
    s_Sql = s_Sql & "IFNULL(cfg.password_envio, '') AS password_envio, IFNULL(cfg.correo_envio, '') AS correo_envio, "
    s_Sql = s_Sql & "IFNULL(cfg.puerto_envio, 0) AS puerto_envio "
    s_Sql = s_Sql & "FROM plparametroafp prm "
    s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON prm.pdoano=cfg.pdoano "
    s_Sql = s_Sql & "LEFT JOIN plcargo ON cfg." & s_Expresion & "cargo=plcargo.codcgo "
    s_Sql = s_Sql & "INNER JOIN plclasplan cls ON cls.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "WHERE prm.pdoano='" & ps_Ano & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!cpcbasico) = "" Then Beep: MsgBox "Debe configurar los parametro de Básico", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!representante) = ", " Then Beep: MsgBox "Debe configurar el parametro de Representante Legal", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!regpatronal) = "" Then Beep: MsgBox "Debe configurar el parametro Regimen Patronal", vbCritical: Exit Sub
    s_RegPatronal = gdl_Funcion.aTexto(porstRecordset!regpatronal)
    s_FmtBoleta = Trim(IIf(IsNull(porstRecordset!fmtboleta), s_Estado_Ina, porstRecordset!fmtboleta))
    s_Representante = IIf(porstRecordset!repimpbol = s_Estado_Act, gdl_Funcion.aTexto(porstRecordset!representante), IIf(s_FmtBoleta = 6, "Empleador o Representante", ps_NomEmpresa))
    s_Cargo = Trim(IIf(IsNull(porstRecordset!descgo), s_Estado_Ina, porstRecordset!descgo))
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Proceso = "rptbolpag" & nTabIndex
    n_ServerEnvio = porstRecordset!server_envio
    s_UsuarioEnvio = porstRecordset!usuario_envio
    s_PasswordEnvio = gdl_Funcion.Desencripta(porstRecordset!password_envio)
    s_CorreoEnvio = porstRecordset!correo_envio
    n_PuertoEnvio = porstRecordset!puerto_envio
    s_TmpCarpeta = "c:\tmpexportar\"
    
    ' valido envio de correo
    If Index = 10 Then
      s_OldMessage = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, ps_Usuario, "PW")
      If (s_OldMessage = "???" Or s_OldMessage = "") Then Beep: MsgBox "Usuario No Registrado", vbExclamation: Exit Sub
      Inputbox_Password fReporteBoleta
      s_Expresion = InputBox("Ingrese Clave de Usuario Acceso Proceso " & Trim(cmdAction(Index).ToolTipText), "Clave de Acceso")
      If UCase(s_Expresion) <> UCase(s_OldMessage) Then Beep: MsgBox "Clave de Usuario Acceso Proceso " & Trim(cmdAction(Index).ToolTipText) & " No es Correcta", vbExclamation: Exit Sub
      If StrConv(dir$(s_TmpCarpeta, vbDirectory), vbLowerCase) = vbNullString Then Beep: MsgBox "No existe Directorio de Proceso '" & s_TmpCarpeta & "'; Verificar", vbCritical: Exit Sub
      ribCopia.Value = False
    End If
            
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Información ..."
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
      tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Realizo la impresión de los registros
    If Index = 8 Then
      ' Genera la información del reporte
      PrnBoletaFormato nTabIndex, s_Proceso, s_FechaHora, IIf(ribCopia.Value, 2, 1)
      'Libera Documentos en la Impresora
      Printer.EndDoc
    Else
      ' Parametros de Impresión
      gdl_Procedure.ps_ReportTitle = Me.Caption
      gdl_Procedure.ps_ReportName = "rptbolpago" & s_FmtBoleta
      ReDim aElemento(3, 7): ReDim aElementos(2)
      ' Parametros del store procedure
      aElemento(0, 0) = ps_CodEmpresa
      aElemento(0, 1) = "": aElemento(0, 2) = ""
      aElemento(0, 3) = "":  aElemento(0, 4) = ""
      ' Formulas del Reporte
      aElemento(1, 0) = "": aElemento(1, 1) = ""
      aElemento(1, 2) = "":  aElemento(1, 3) = "":  aElemento(1, 4) = ""
      ' Campos de Parametros del Reporte
      aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
      aElemento(2, 1) = "TituloReporte;" & IIf(s_FmtBoleta = "3", "RECIBO DE PAGO (", "BOLETA DE PAGO (") & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")" & ";true"
      aElemento(2, 2) = "Ruc;" & ps_RucEmpresa & ";true"
      aElemento(2, 3) = "RegPatronal;" & s_RegPatronal & ";true"
      aElemento(2, 4) = "Representante;" & s_Representante & ";true"
      aElemento(2, 5) = "Direccion;" & sDireccion & ";true"
      aElemento(2, 6) = "Cargo;" & s_Cargo & ";true"
      
      ' Filtro de Formulas y Grupos del Reporte
      aElementos(0) = ""
      aElementos(1) = ""
      
      ' [ Generación e impresión de información para el reporte
      s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
      gdl_Conexion.Execucion s_Sql, Elimina
      
      s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
      s_Sql = s_Sql & "codpsn varchar(11) Null, nompsn varchar(80) Null, "
      s_Sql = s_Sql & "codcgo char(3) Null, descgo varchar(80) Null, "
      s_Sql = s_Sql & "fecingreso date Null, nroessalud varchar(15) Null, "
      s_Sql = s_Sql & "despdo varchar(40) Null, desconfianza varchar(50) Null, "
      s_Sql = s_Sql & "fechainivacacion date Null, fechafinvacacion date Null, numdociden varchar(11) Null, "
      s_Sql = s_Sql & "numeroafp varchar(15) Null, codafp char(2) Null, desafp varchar(40) Null, "
      s_Sql = s_Sql & "diatrabajo decimal(6,2) Null Default '0', diaferiado decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "horanormal decimal(6,2) Null Default '0', horaextra decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "horaextra1 decimal(6,2) Null Default '0', horaextra2 decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "horaextra3 decimal(6,2) Null Default '0', horaextra4 decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "horaextrax decimal(6,2) Null Default '0', fecbaja date Null, copia char(1) Null, "
      s_Sql = s_Sql & "secuencia smallint(3) Null, basico decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcing varchar(4) Null, descpcing varchar(40) Null, impcpcing decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcdsc varchar(4) Null, descpcdsc varchar(40) Null, impcpcdsc decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcapo varchar(4) Null, descpcapo varchar(40) Null, impcpcapo decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "moneda char(3) Null, monpago char(3) Null, importipcmb decimal(6,3) Null Default '0', impornetocmb decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "salinicte decimal(18,2) Null Default '0', salfincte decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "detcco varchar(60) Null, descdt varchar(60) Null, "
      s_Sql = s_Sql & "fechaini date Null, fechafin date Null, "
      s_Sql = s_Sql & "dialaboral decimal(6,2) Null Default '0', dianolaboral decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "diasubsidio decimal(6,2) Null Default '0', diafalta decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "diavacaciones decimal(6,2) Null Default '0', diatradesemanal decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "diasuspension decimal(6,2) Null Default '0', dialibre decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "dialicencia decimal(6,2) Null Default '0', diaenfermedad decimal(6,2) Null Default '0', "
      s_Sql = s_Sql & "tipdoc varchar(40) Null, tiptra varchar(40) Null, "
      s_Sql = s_Sql & "tardanza decimal(6,2) Null Default '0', seccion varchar(40) Null, ubicacion varchar(40) Null, codtab varchar(11) Null, "
      s_Sql = s_Sql & "PRIMARY KEY (codtab, codpsn, copia, secuencia))"
      gdl_Conexion.Execucion s_Sql, Inserta
      ' Genera la información del reporte
      PrnBoletaGenerica nTabIndex, s_Proceso, s_FechaHora, "tmp" & gdl_Procedure.ps_ReportName, Index
      s_Expresion = IIf(ribFirma.Value, "nexo", "")
      
      ' Selecciono información del reporte
      s_Sql = "SELECT bol.codpsn, bol.nompsn, bol.codcgo, bol.descgo, "
      s_Sql = s_Sql & "bol.fecingreso, bol.nroessalud, bol.despdo, bol.desconfianza, "
      s_Sql = s_Sql & "bol.fechainivacacion, bol.fechafinvacacion, bol.numdociden, "
      s_Sql = s_Sql & "bol.numeroafp, bol.codafp, bol.desafp, bol.diatrabajo, bol.horanormal, "
      s_Sql = s_Sql & "bol.horaextra, bol.horaextra1, bol.horaextra2, bol.horaextra3, "
      s_Sql = s_Sql & "bol.horaextrax, bol.fecbaja, bol.copia, bol.secuencia, bol.basico, "
      s_Sql = s_Sql & "bol.codcpcing, bol.descpcing, bol.impcpcing, "
      s_Sql = s_Sql & "bol.codcpcdsc, bol.descpcdsc, bol.impcpcdsc, "
      s_Sql = s_Sql & "bol.codcpcapo, bol.descpcapo, bol.impcpcapo, "
      s_Sql = s_Sql & "bol.moneda, bol.monpago, bol.importipcmb, bol.impornetocmb, "
      s_Sql = s_Sql & "bol.salinicte, bol.salfincte, bol.detcco, "
      s_Sql = s_Sql & "bol.fechaini, bol.fechafin, "
      If s_FmtBoleta = 6 Then
        s_Sql = s_Sql & "psn.fecnacimiento, "
        s_Sql = s_Sql & "CONCAT(IFNULL(via.abrevia,''), ' ', IFNULL(psn.nomviadirec,''), ' ', IFNULL(psn.numerdirec, ''), ' ', IFNULL(psn.intedirec, ''), ' ', IFNULL(zon.abrezona, ''), ' ', IFNULL(psn.nomzondirec, ''), ' ', IFNULL(ubg.desubg, '')) AS psndireccion, "
        s_Sql = s_Sql & "bol.descdt, "
      ElseIf s_FmtBoleta = 8 Then
        s_Sql = s_Sql & "bol.dialaboral, bol.diaferiado, bol.diasubsidio, bol.diafalta, bol.dianolaboral, bol.horaextra4, "
        s_Sql = s_Sql & "bol.tiptra, bco.desbco, psn.cuentapago, pdo.fechaproceso AS fechapago, bol.codtab, dci.sigladci, bol.diaenfermedad, "
      ElseIf s_FmtBoleta = 10 Then
        s_Sql = s_Sql & "bol.tipdoc, bol.tiptra, "
      ElseIf s_FmtBoleta = 12 Then
        s_Sql = s_Sql & "bol.diasubsidio, bol.diafalta, bol.tardanza, "
      ElseIf s_FmtBoleta = 13 Then
        s_Sql = s_Sql & "bol.dialaboral, bol.diaferiado, bol.diasubsidio, bol.diafalta, bol.dianolaboral, bol.horaextra4, "
        s_Sql = s_Sql & "bol.tiptra, bco.desbco, psn.cuentapago, pdo.fechaproceso AS fechapago, bol.codtab, dci.sigladci, "
        s_Sql = s_Sql & "bol.diavacaciones, bol.diatradesemanal, bol.diasuspension, bol.dialibre, bol.dialicencia, bol.diaenfermedad, "
        s_Sql = s_Sql & "bol.ubicacion, tco.destco, "
      End If
      'Enero 2015
      's_Sql = s_Sql & "cfg.logo, cfg.firma" & s_Expresion & ", " & IIf(ribOrdenar.Value, "bol.nompsn", "bol.codpsn") & " AS bolorden "
       s_Sql = s_Sql & "0 as logo, 0 as firma" & s_Expresion & ", " & IIf(ribOrdenar.Value, "bol.nompsn", "bol.codpsn") & " AS bolorden "
       s_Sql = s_Sql & "FROM (tmp" & gdl_Procedure.ps_ReportName & " bol, plcfgempresa cfg) "
       ' s_Sql = s_Sql & "FROM (tmp" & gdl_Procedure.ps_ReportName & " bol) "
      
     
      s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON psn.codpsn=bol.codpsn AND psn.codcls='" & ps_ClsPlanilla & "' "
      
      If s_FmtBoleta = 6 Then
        s_Sql = s_Sql & "LEFT JOIN pltipovia via ON via.codvia=psn.codvia "
        s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON zon.codzona=psn.codzona "
        s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON ubg.codubg=psn.ubigeodir "
      ElseIf s_FmtBoleta = 8 Then
        s_Sql = s_Sql & "LEFT JOIN plbanco bco ON bco.codbco=psn.codbcopago "
        s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON dci.coddci=psn.coddci "
        s_Sql = s_Sql & "LEFT JOIN plperiodo pdo ON pdo.codcls='" & ps_ClsPlanilla & "' AND pdo.codpdo='" & txtPeriodo.Text & "' "
      ElseIf s_FmtBoleta = 13 Then
        s_Sql = s_Sql & "LEFT JOIN plbanco bco ON bco.codbco=psn.codbcopago "
        s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON dci.coddci=psn.coddci "
        s_Sql = s_Sql & "LEFT JOIN plperiodo pdo ON pdo.codcls='" & ps_ClsPlanilla & "' AND pdo.codpdo='" & txtPeriodo.Text & "' "
        s_Sql = s_Sql & "LEFT JOIN plcontrato con ON con.codcls='" & ps_ClsPlanilla & "' AND con.codpsn=bol.codpsn AND con.estadocon='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "AND con.fechafin =(SELECT MAX(cmx.fechafin) FROM plcontrato cmx WHERE cmx.codcls='" & ps_ClsPlanilla & "' AND cmx.codpsn=bol.codpsn AND cmx.estadocon='" & s_Estado_Act & "') "
        s_Sql = s_Sql & "LEFT JOIN pltipcontrato tco ON tco.codtco=con.tipcon "
      End If
      
      s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Ano & "' "
      s_PdfWhere = "ORDER BY codtab, " & IIf(ribOrdenar.Value, "nompsn", "codpsn") & ", copia, secuencia"
      
      ' Información archivo pdf
      If Index >= 9 Then
        s_PdfWhere = "SELECT DISTINCT bol.codpsn, bol.numdociden, psn.correoelect "
        s_PdfWhere = s_PdfWhere & "FROM tmp" & gdl_Procedure.ps_ReportName & " bol "
        s_PdfWhere = s_PdfWhere & "INNER JOIN plpersonal psn ON bol.codpsn=psn.codpsn AND psn.codcls='" & ps_ClsPlanilla & "' "
        s_PdfWhere = s_PdfWhere & "ORDER BY codpsn"
        Set porstExporta = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_PdfWhere)
        If porstExporta.RecordCount > 0 Then
          
             If Index = 9 Then
                MuestraMensaje "Generando Boletas en Archivos PDF ..."
             Else
                MuestraMensaje "Envio de Información ..."
             End If

          ' Instancio objeto pdfexporta
          Set oCreadorPdf = CreateObject("syslink.creadorpdf")
          oCreadorPdf.o_cStart
          Set fso = CreateObject("Scripting.FileSystemObject") 'enero 2015
          
          
          While Not porstExporta.EOF
            s_PdfWhere = "AND bol.codpsn='" & porstExporta!codpsn & "' "
            s_PdfWhere = s_PdfWhere & "ORDER BY codtab, " & IIf(ribOrdenar.Value, "nompsn", "codpsn") & ", copia, secuencia"
            Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql & s_PdfWhere)
            
            s_Expresion = s_TmpCarpeta & porstExporta!numdociden & ".pdf"
            ' Verifico exista Archivo boleta
           nTemporizador = 0
            
            Do While (StrConv(dir$(s_Expresion, vbHidden), vbLowerCase) = vbNullString)
              oCreadorPdf.o_cConfigPrint s_TmpCarpeta, porstExporta!numdociden
              ' Ejecuto reporte y saco de memoria la información
              gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, s_Estado_Act, False, True, False, True, True, aElemento, aElementos, porstRecordset
              'Si despues de 3 intentos no se ha podido crear el pdf, deja de intentar para pasar a leer el siguiente registro.
              If nTemporizador = 3 Then
                Exit Do
              End If
              nTemporizador = nTemporizador + 1
            Loop
            
            ' correo adjunto boleta
            If Index = 10 Then
              ' envio correo electronico
              If n_ServerEnvio = s_Estado_Ina Then
                EnviaCorreoOutlook porstExporta!correoelect, "Boleta de pago " & lblHelp(1).Caption, "Buen día adjunto boleta de pago correspondiente a " & lblHelp(1).Caption & "; confirmar recepción de este correo", "", "", s_Expresion, 1
              Else
                EnviaCorreoCDOWeb n_ServerEnvio, s_UsuarioEnvio, s_PasswordEnvio, s_CorreoEnvio, porstExporta!correoelect, "Boleta de pago " & lblHelp(1).Caption, "Buen día adjunto boleta de pago correspondiente a " & lblHelp(1).Caption & "; confirmar recepción de este correo", "", "", s_Expresion, n_PuertoEnvio, 1
              End If
            End If
            
            porstExporta.MoveNext
          
          Wend
          Set oCreadorPdf = Nothing
          Set porstExporta = Nothing
        End If
      Else
        Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql & s_PdfWhere)
        ' Ejecuto reporte y saco de memoria la información
        gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
        
        'Enero 2015
        
         FileCopy "C:\Program Files (x86)\personal y planilla\bmp\mifirma.bmp", "C:\tmpexportar\firma.bmp"
         FileCopy "C:\Program Files (x86)\personal y planilla\bmp\milogo.bmp", "C:\tmpexportar\logo.bmp"
         
      End If
      Set porstRecordset = Nothing
      ' Elimino la tabla temporal de impresion
      s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
      gdl_Conexion.Execucion s_Sql, Elimina
    End If
    ' Elimino el rango de impresion
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, "", ps_Usuario, s_FechaHora, "E"
    
    ' Elimino archivos temporales envio correo
    If Index = 10 Then
        s_Expresion = s_TmpCarpeta & "*.pdf"
        If Not dir$(s_Expresion) = vbNullString Then
          Kill s_Expresion
        End If
    End If
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
    '[ Finalizo la conexión a la base de datos ]
    Set gdl_Conexion = Nothing
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim s_TablaHelp As String
  
  s_SqlHelp = ""
  If n_IndexHelp = Index And Index <> 1 Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If

  Select Case Index
   Case 0     ' Formato de Boleta
    tdbHelp.Columns(0).DataField = "codboleta": tdbHelp.Columns(1).DataField = "desboleta"
    s_TablaHelp = "Formato de Boletas"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("bol", "codboleta", ps_ClsPlanilla, "")
   Case 1     ' Periodo de Pago
    tdbHelp.Columns(0).DataField = "codpdo": tdbHelp.Columns(1).DataField = "despdo"
    s_TablaHelp = "Periodos de Pago"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ped", "codpdo", s_Estado_Ina & ps_ClsPlanilla & ps_Ano, "")
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
  Me.Height = 7110: Me.Width = 8000
  Me.Left = 120: Me.Top = 80
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Trabajador(es)"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Apellido Paterno", "Apellido Materno", "Nombre(s)", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "apepaterno", "apematerno", "nombres", "estadopsn")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1080, 1616.33, 1616.33, 1616.33, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "")
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
  gdl_Procedure.InicializaGrilla tdbSeleccion(0), aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbSeleccion(0).Columns(4).ValueItems.Presentation = dbgNormal
  tdbSeleccion(0).Columns(4).ValueItems.Translate = True
  For n_Index = 0 To 5
    tdbSeleccion(0).Columns(4).ValueItems.Add Item
    tdbSeleccion(0).Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "A", "V", "L", "P", "O", "I")
    tdbSeleccion(0).Columns(4).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadovo", "estadnok", "estadopk", "estadopn", "procenok") & ".bmp")
  Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbSeleccion(0), s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbSeleccion(0).GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbSeleccion(0).AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(11, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "imprimir", "imprapid", "consolid", "genemail")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir", "Imprime Formato", "Archivo PDF", "Correo Electrónico")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' grafico boton orden alfabetico
  ribOrdenar.PictureUp = LoadPicture()
  ribOrdenar.ToolTipText = "Reporte Alfabeticamente"
  s_Sql = gdl_Procedure.ps_PathImagen & "ordalfab.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribOrdenar.PictureUp = LoadPicture(s_Sql)
  ribOrdenar.Value = False
  ' Cargo el grafico del boton de copia
  ribCopia.PictureUp = LoadPicture()
  ribCopia.ToolTipText = "Copia de Boleta "
  s_Sql = gdl_Procedure.ps_PathImagen & "actsaldo.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribCopia.PictureUp = LoadPicture(s_Sql)
  ribCopia.Value = False
  ' Cargo el grafico del boton de firma
  ribFirma.PictureUp = LoadPicture()
  ribFirma.ToolTipText = "Representante Adjunto"
  s_Sql = gdl_Procedure.ps_PathImagen & "dividir.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribFirma.PictureUp = LoadPicture(s_Sql)
  ribFirma.Value = False
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codbco", "desbco")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 834.7402, 3345.071)
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Periodo de Pago", 2
  ']
  
  '[ Configuro las grillas de selección
  ReDim aElemento(3, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcco", "detcco", "estcco")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1000, 3555.213, 300)
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
  For n_Index = 1 To 3
    aElemento(0, 1) = Choose(n_Index, "codcco", "codubica", "codsec")
    aElemento(1, 1) = Choose(n_Index, "detcco", "desubica", "dessec")
    aElemento(2, 1) = Choose(n_Index, "estcco", "estadoubica", "estadosec")
    gdl_Procedure.InicializaGrilla tdbSeleccion(n_Index), aElemento, aElementos
    ' Cambio el formato de la grilla columna de valores
    tdbSeleccion(n_Index).Columns(2).ValueItems.Presentation = dbgNormal
    tdbSeleccion(n_Index).Columns(2).ValueItems.Translate = True
    ' Primera columna
    tdbSeleccion(n_Index).Columns(2).ValueItems.Add Item
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(0).Value = IIf(n_Index = 1, "A", s_Estado_Act)
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(0).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & "estadok.bmp")
    ' Segunda columna
    tdbSeleccion(n_Index).Columns(2).ValueItems.Add Item
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(1).Value = IIf(n_Index = 1, "I", s_Estado_Ina)
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(1).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & "estadnok.bmp")
    ' Personaliza el estilo de la grilla de TDBGrid
    gdl_Procedure.DefineStyleGrilla tdbSeleccion(n_Index), Choose(n_Index, "Centro Costo", "Ubicación o Localidad", "Sección de Empresa"), 1
    ' Agrupacion de columnas y titulo DataView = dbgGroupView
    tdbSeleccion(n_Index).GroupByCaption = "Arrastrar titulo de columna de agrupación"
  Next n_Index
  ']
  
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  For n_Index = 0 To 3
    tdbSeleccion(n_Index).DataSource = dcaSeleccion(n_Index)
    If n_Index <> 0 Then RecuperaRegistros n_Index, tdbSeleccion(n_Index).Columns(0).DataField & " ASC"
  Next n_Index
  ribParametro(0).Value = True
  
  'FEBRERO 2015
  'Deshabiltio Panel 2 por no estar terminado
  
  panTool(1).Enabled = False
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

Private Sub ribParametro_Click(Index As Integer, Value As Integer)
  RecuperaRegistros 0, tdbSeleccion(0).Columns(0).DataField & " ASC"
End Sub

Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Formato de boletas de pago
    txtFormato = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtFormato.SetFocus
   Case 1       ' Periodo de pago
    txtPeriodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo.SetFocus
  End Select
   
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Formatos de boleta
    s_Sql = gdl_Funcion.HelpTablas("bol", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
   Case 1     ' Periodo de Pago
    s_Sql = gdl_Funcion.HelpTablas("ped", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & ps_Ano, "")
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
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaSeleccion(Index), tdbSeleccion(Index), " " & tdbSeleccion(Index).Caption
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
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtFormato, "BP")
End Sub
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, "PR")
End Sub

