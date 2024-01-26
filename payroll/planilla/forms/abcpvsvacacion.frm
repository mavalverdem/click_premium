VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcPvsVacacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4110
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "abcpvsvacacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   7740
   Begin TabDlg.SSTab tabRegister 
      Height          =   2910
      Left            =   75
      TabIndex        =   32
      Top             =   600
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5133
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
      TabPicture(0)   =   "abcpvsvacacion.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "shpCuadro(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblHelp(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDato(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdHelp(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmCuadro(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frmCuadro(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCodigo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtPeriodo(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtNumero"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPeriodo(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.TextBox txtPeriodo 
         Height          =   300
         Index           =   1
         Left            =   4245
         TabIndex        =   6
         Top             =   555
         Width           =   980
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Left            =   3390
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1740
         Width           =   800
      End
      Begin VB.TextBox txtPeriodo 
         Height          =   300
         Index           =   0
         Left            =   1340
         TabIndex        =   4
         Top             =   555
         Width           =   980
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1340
         TabIndex        =   1
         Top             =   210
         Width           =   980
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1170
         Index           =   1
         Left            =   4530
         TabIndex        =   14
         Top             =   1050
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   2064
         _StockProps     =   14
         Caption         =   " Estado "
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
         Begin Threed.SSOption optEstado 
            Height          =   190
            Index           =   0
            Left            =   180
            TabIndex        =   15
            Top             =   285
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   335
            _StockProps     =   78
            Caption         =   "&Pendiente"
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
         End
         Begin Threed.SSOption optEstado 
            Height          =   190
            Index           =   1
            Left            =   180
            TabIndex        =   16
            Top             =   570
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   335
            _StockProps     =   78
            Caption         =   "P&rovisionado"
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
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   17
            Top             =   840
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   335
            _StockProps     =   78
            Caption         =   "&Cancelado"
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
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1050
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   1125
         Width           =   2940
         _Version        =   65536
         _ExtentX        =   5186
         _ExtentY        =   1852
         _StockProps     =   14
         Caption         =   " Fecha "
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
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   9
            Top             =   615
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            Format          =   141230081
            CurrentDate     =   37515
         End
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   1
            Left            =   1530
            TabIndex        =   11
            Top             =   615
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   141230081
            CurrentDate     =   37515
         End
         Begin VB.Label lblDato 
            Caption         =   "Final :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   1530
            TabIndex        =   10
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Inicio :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   8
            Top             =   330
            Width           =   1005
         End
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   2385
         TabIndex        =   33
         Top             =   210
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Final :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3090
         TabIndex        =   5
         Top             =   600
         Width           =   1005
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
         Left            =   2745
         TabIndex        =   2
         Top             =   255
         Width           =   195
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Días :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3390
         TabIndex        =   12
         Top             =   1440
         Width           =   795
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1350
         Index           =   0
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   975
         Width           =   4140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicial :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio :"
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
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   255
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   18
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
      Begin Threed.SSCommand cmdCancel 
         Height          =   360
         Left            =   6690
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
         Picture         =   "abcpvsvacacion.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6300
         TabIndex        =   20
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
         Picture         =   "abcpvsvacacion.frx":0044
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
         TabIndex        =   21
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   3600
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
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   3
         Left            =   4935
         TabIndex        =   23
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
         Picture         =   "abcpvsvacacion.frx":0060
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4545
         TabIndex        =   24
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
         Picture         =   "abcpvsvacacion.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2835
         TabIndex        =   25
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
         Picture         =   "abcpvsvacacion.frx":0098
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2445
         TabIndex        =   26
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
         Picture         =   "abcpvsvacacion.frx":00B4
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   2910
      Index           =   0
      Left            =   6960
      TabIndex        =   27
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   5133
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
         TabIndex        =   28
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
         TabIndex        =   29
         Tag             =   "0"
         Top             =   585
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcpvsvacacion.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   30
         Tag             =   "0"
         Top             =   1395
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcpvsvacacion.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   31
         Tag             =   "0"
         Top             =   2175
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcpvsvacacion.frx":0108
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2625
      TabIndex        =   34
      Top             =   420
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
Attribute VB_Name = "fAbcPvsVacacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_Registro As String                            ' Codigo del registro
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins)

End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fPvsVacacion.dcaRegistro.Recordset!codpvs.DefinedSize
    gdl_Procedure.EditText "PK", txtPeriodo(0), "", s_MdoData_Upd, False, 4
    gdl_Procedure.EditText "PK", txtPeriodo(1), "", s_MdoData_Upd, False, 4
    gdl_Procedure.EditText "AT", txtnumero, 0, Me.Tag, False, 4, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(2), False, Me.Tag, False
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fPvsVacacion.dcaRegistro.Recordset!codpvs, Me.Tag, True, fPvsVacacion.dcaRegistro.Recordset!codpvs.DefinedSize
    gdl_Procedure.EditText "PK", txtPeriodo(0), Left(gdl_Funcion.aTexto(fPvsVacacion.dcaRegistro.Recordset!pdopvs), 4), Me.Tag, False, 4
    gdl_Procedure.EditText "PK", txtPeriodo(1), Mid(gdl_Funcion.aTexto(fPvsVacacion.dcaRegistro.Recordset!pdopvs), 5), Me.Tag, False, 4
    gdl_Procedure.EditText "AT", txtnumero, CInt(fPvsVacacion.dcaRegistro.Recordset!numerodias), Me.Tag, False, 4, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), fPvsVacacion.dcaRegistro.Recordset!fechaini, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), fPvsVacacion.dcaRegistro.Recordset!fechafin, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fPvsVacacion.dcaRegistro.Recordset!estadovac = s_Estado_Ina), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fPvsVacacion.dcaRegistro.Recordset!estadovac = s_Estado_Act), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(2), (fPvsVacacion.dcaRegistro.Recordset!estadovac = s_Estado_Blq), Me.Tag, False
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtCodigo.Text, "PV")

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_Periodo As String
  
  ' Valido que el periodo no se eencuentre procesado
  If optEstado(1).Value And Index <> 0 Then Beep: MsgBox "Sub Periodo No se puede Actualizar se encuentra Provisionado - Cancelado", vbExclamation: Me.Tag = s_MdoData_Vis: Exit Sub
  If Not (o_PvsVacaciones.dcaRegistro.Recordset!estadopsn <> "I" And o_PvsVacaciones.dcaRegistro.Recordset!remintegralvaca = s_Estado_Ina) And Index <> 0 Then Beep: MsgBox "Personal se encuentra inactivo o tiene vacaciones integral", vbExclamation: Me.Tag = s_MdoData_Vis: Exit Sub
  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   txtPeriodo(0).SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & lblHelp(0) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim(txtCodigo.Text)
    s_Periodo = Trim(txtPeriodo(0).Text) & Trim(txtPeriodo(1).Text)
    
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpvs", "codpsn", "pdopvs")
    a_Valores = Array(ps_ClsPlanilla, s_Registro, o_PvsVacaciones.dcaRegistro.Recordset!codpsn, s_Periodo)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plpvsvacacion", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el ado control y la grilla
    gdl_Procedure.RefreshAdoControl fPvsVacacion.dcaRegistro, fPvsVacacion.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fPvsVacacion.dcaRegistro.Recordset.EOF And fPvsVacacion.dcaRegistro.Recordset.BOF) Or fPvsVacacion.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fPvsVacacion.dcaRegistro.Recordset.Find ("cPrimaryKey='" & s_Registro & s_Periodo & "'")
      If fPvsVacacion.dcaRegistro.Recordset.EOF Then fPvsVacacion.dcaRegistro.Recordset.MoveLast
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
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Ejercicio de provisión - fisicas
    tdbHelp.Columns(0).DataField = "codpvs": tdbHelp.Columns(1).DataField = "descripvs"
    tdbHelp.Caption = "Periodos de Provisión"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("vxe", "codpvs", s_Estado_Ina & ps_ClsPlanilla, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Left = (tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero inicial, anterior, siguiente o final
  Select Case Index
   Case 0: fPvsVacacion.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fPvsVacacion.dcaRegistro.Recordset.BOF Then fPvsVacacion.dcaRegistro.Recordset.MovePrevious
           If fPvsVacacion.dcaRegistro.Recordset.BOF Then fPvsVacacion.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fPvsVacacion.dcaRegistro.Recordset.EOF Then fPvsVacacion.dcaRegistro.Recordset.MoveNext
           If fPvsVacacion.dcaRegistro.Recordset.EOF Then fPvsVacacion.dcaRegistro.Recordset.MoveLast
   Case 3: fPvsVacacion.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1, s_Periodo As String
  Dim dFecFinal As Date
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If Not IsNumeric(txtnumero.Text) Or CInt(txtnumero.Text) < 0 Or CInt(txtnumero.Text) > 360 Then Beep: MsgBox "Numero de días  no valido; verifique", vbExclamation: txtnumero.SetFocus: Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha final debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  If Not (o_PvsVacaciones.dcaRegistro.Recordset!estadopsn <> "I" And o_PvsVacaciones.dcaRegistro.Recordset!remintegralvaca = s_Estado_Ina) Then Beep: MsgBox "Personal se encuentra inactivo o tiene vacaciones integrales", vbExclamation: txtCodigo.SetFocus: Exit Sub
  If Not (Trim(txtCodigo.Text) > Format(o_PvsVacaciones.dcaRegistro.Recordset!fecingreso, "yyyy")) Then Beep: MsgBox "Ejercicio no corresponde al personal", vbExclamation: txtCodigo.SetFocus: Exit Sub
  If Not (Format(dtpFechas(0), "yyyymmdd") >= Format(o_PvsVacaciones.dcaRegistro.Recordset!fecingreso, "yyyymmdd")) Then Beep: MsgBox "Fecha Inicial debe ser mayor o igual que la fecha de Ingreso", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  If Not (Format(dtpFechas(0), "yyyy") < txtCodigo.Text) Then Beep: MsgBox "Fecha Inicial debe ser menor del periodo de Provisión", vbExclamation: dtpFechas(0).SetFocus: Exit Sub
  If Not (Format(dtpFechas(1), "yyyy") <= txtCodigo.Text) Then Beep: MsgBox "Fecha Final debe ser igual del periodo de Provisión", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  dFecFinal = DateAdd("yyyy", 1, dtpFechas(0)) - 1
  If Not (dtpFechas(1) <= dFecFinal) Then Beep: MsgBox "Fecha Final debe ser menor o igual que '" & dFecFinal & "'", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, s_Estado_Ina, s_Estado_Act)
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(txtCodigo.Text)
  s_Periodo = Trim(txtPeriodo(0).Text) & Trim(txtPeriodo(1).Text)
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpvs", "codpsn", "pdopvs", "fechaini", "fechafin", "numerodias", "estadovac", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, s_Registro, o_PvsVacaciones.dcaRegistro.Recordset!codpsn, s_Periodo, Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), CLng(txtnumero.Text), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpvs", "codpsn", "pdopvs")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plpvsvacacion", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plpvsvacacion", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fPvsVacacion.dcaRegistro, fPvsVacacion.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fPvsVacacion.dcaRegistro.Recordset.Find ("cPrimaryKey='" & s_Registro & s_Periodo & "'")
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
Private Sub dtpFechas_LostFocus(Index As Integer)
  Dim n_Dias  As Long
  
  If Index = 0 Then Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha final debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  n_Dias = gdl_Funcion.NumeroDias360(dtpFechas(1), dtpFechas(0), dtpFechas(1))
  txtnumero.Text = CLng(n_Dias)

End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 4590: Me.Width = 7830
  Me.Left = 2300: Me.Top = 1800
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Sub Periodos de Provisión"
  lblTitle = "Sub Periodo de Provisión"
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fPvsVacacion.Tag

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
  l_ExistRecord = (fPvsVacacion.dcaRegistro.Recordset.EOF Or fPvsVacacion.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fPvsVacacion.dcaRegistro.Recordset!pdopvs
  
  ' Carga los datos en el formulario
  ShowScreen
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcgo", "descgo")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Cargos de Personal", 2
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Ejercicion de provsión - fisicas
    txtCodigo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCodigo.SetFocus
  End Select
  
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0  ' Ejercicio de provisión fisicas
    s_Sql = gdl_Funcion.HelpTablas("vxe", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla, "")
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
      SendKeys "{TAB}"
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txtCodigo_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtCodigo, "PV")
  txtPeriodo(0).Text = "": txtPeriodo(1).Text = ""
  If (lblHelp(0) = "???" Or lblHelp(0) = "") Then Exit Sub
  txtPeriodo(1).Text = txtCodigo.Text
  txtPeriodo(0).Text = Val(txtPeriodo(1).Text) - 1
End Sub
Private Sub txtnumero_GotFocus()
  gdl_Procedure.MarcaGet txtnumero
End Sub
Private Sub txtnumero_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNumero_Validate(Cancel As Boolean)
  txtnumero.Text = IIf(Not IsNumeric(txtnumero.Text), 0, txtnumero.Text)
  txtnumero.Text = IIf(CInt(txtnumero.Text) < 0, 0, txtnumero.Text)
  txtnumero.Text = IIf(CInt(txtnumero.Text) > 360, 360, txtnumero.Text)
  txtnumero.Text = CInt(txtnumero.Text)
End Sub
Private Sub txtPeriodo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPeriodo(Index)
End Sub
Private Sub txtPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
