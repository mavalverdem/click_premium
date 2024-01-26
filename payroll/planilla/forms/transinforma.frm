VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fTransInformacio 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6420
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6645
   Icon            =   "transinforma.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   6645
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   1335
      Left            =   6960
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2355
      _LayoutType     =   0
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
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
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
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=13,.bold=0,.fontsize=825,.italic=0"
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
   Begin MSComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   6050
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin Threed.SSFrame sfmProgreso 
      Height          =   550
      Left            =   75
      TabIndex        =   28
      Top             =   5790
      Width           =   6480
      _Version        =   65536
      _ExtentX        =   11430
      _ExtentY        =   970
      _StockProps     =   14
      Caption         =   " Procesando archivo : "
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
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   5205
      Left            =   75
      TabIndex        =   25
      Top             =   555
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   9181
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
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
      TabCaption(0)   =   "Importación"
      TabPicture(0)   =   "transinforma.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmCuadro(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmCuadro(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmCuadro(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.OptionButton Option2 
         Caption         =   "Desmarcar Todo"
         Height          =   195
         Left            =   4200
         TabIndex        =   40
         Top             =   5040
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Marcar Todo"
         Height          =   195
         Left            =   2300
         TabIndex        =   39
         Top             =   5040
         Value           =   -1  'True
         Width           =   1400
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   4515
         Index           =   1
         Left            =   3870
         TabIndex        =   17
         Top             =   60
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   7964
         _StockProps     =   14
         Caption         =   " Ubicación "
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
         Begin VB.FileListBox flbArchivo 
            Height          =   1845
            Index           =   0
            Left            =   135
            Pattern         =   "*.sma"
            TabIndex        =   22
            Top             =   2550
            Width           =   2235
         End
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Index           =   0
            Left            =   135
            TabIndex        =   20
            Top             =   795
            Width           =   2235
         End
         Begin VB.DriveListBox drbUnidad 
            Height          =   315
            Index           =   0
            Left            =   140
            TabIndex        =   19
            Top             =   495
            Width           =   2240
         End
         Begin VB.Label lblDato 
            Caption         =   "Archivos :"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   21
            Top             =   2325
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   140
            TabIndex        =   18
            Top             =   250
            Width           =   1005
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3195
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   60
         Width           =   3645
         _Version        =   65536
         _ExtentX        =   6429
         _ExtentY        =   5636
         _StockProps     =   14
         Caption         =   " Tablas "
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
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   255
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3351
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Entidad Banacaria "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   495
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3351
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Entidad de Pensión"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   750
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Entidad Pres. Salud"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   4
            Top             =   990
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Cargos de Personal"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   5
            Top             =   1230
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3351
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Profesión u Oficio"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   6
            Top             =   1485
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3351
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Documento Identidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   7
            Top             =   1725
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Conceptos de Calculo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   8
            Top             =   1965
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Ubicación o Localidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   9
            Top             =   2220
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Sección de empresas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   10
            Top             =   2460
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Centro de Costo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   11
            Top             =   2700
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Padrón de Personal"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkTabla 
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   12
            Top             =   2940
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Proceso de Calculo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1590
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   3240
         Width           =   3660
         _Version        =   65536
         _ExtentX        =   6456
         _ExtentY        =   2805
         _StockProps     =   14
         Caption         =   " Procesos "
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
         Begin VB.TextBox txtPeriodo 
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   1
            Left            =   2200
            MaxLength       =   8
            TabIndex        =   32
            Top             =   735
            Width           =   885
         End
         Begin VB.TextBox txtPeriodo 
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   0
            Left            =   2200
            MaxLength       =   8
            TabIndex        =   29
            Top             =   180
            Width           =   885
         End
         Begin Threed.SSCheck chkProceso 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   285
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Remun Dscto Default "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkProceso 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   540
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Remun Dscto Anterior"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkProceso 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   810
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   353
            _StockProps     =   78
            Caption         =   "Proceso Historico"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   0
            Left            =   3150
            TabIndex        =   30
            Top             =   180
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   1
            Left            =   3150
            TabIndex        =   33
            Top             =   735
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCheck chkProceso 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1065
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Asistencia"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSCheck chkProceso 
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Remun Dscto Excepcional"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   2200
            TabIndex        =   34
            Top             =   1035
            Width           =   135
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   2200
            TabIndex        =   31
            Top             =   480
            Width           =   135
         End
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6645
      _Version        =   65536
      _ExtentX        =   11721
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
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   5940
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
         Picture         =   "transinforma.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   5550
         TabIndex        =   27
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "transinforma.frx":0044
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
         Left            =   390
         TabIndex        =   24
         Top             =   120
         Width           =   4800
      End
   End
End
Attribute VB_Name = "fTransInformacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private s_Registro As String                            ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda

Public Est_CierrePeriodo As String                      ' Estado del cierre del periodo
'[
Private Function ppActualiza_Procesos() As Boolean
  Dim sTabla As String
  Dim nContador As Integer
  Dim nRegistro As Long, nRegistros As Long

  ' Inicializo la barra de progreso
  pgbProgreso.Max = chkProceso.Count
  pgbProgreso.Value = pgbProgreso.Min
  For nContador = 0 To chkProceso.Count - 1
    ' Verifico que se haya seleccionado
    If chkProceso(nContador).Value Then
      sTabla = Choose(nContador + 1, "plremudefa", "plresultado", "plresultado", "plasistencia", "plremuexce")
      sfmProgreso.Caption = " Actualizando Información: " & Trim(chkProceso(nContador).Caption) & " "
      Select Case nContador
       Case 0
        ' Inserto remuneraciones descuentos default
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codpsn, tmp.codcpc, tmp.codmon, tmp.imporemune, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " rxd ON tmp.codcls=rxd.codcls AND tmp.codpsn=rxd.codpsn AND tmp.codcpc=rxd.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(rxd.codcls, rxd.codpsn, rxd.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY codpsn, codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
        
        s_Sql = "UPDATE " & sTabla & " rxd, tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "SET rxd.codmon=tmp.codmon, rxd.imporemune=tmp.imporemune, "
        s_Sql = s_Sql & "rxd.usrmdf=tmp.usrcre, rxd.fyhmdf=tmp.fyhcre "
        s_Sql = s_Sql & "WHERE rxd.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND rxd.codcls=tmp.codcls AND rxd.codpsn=tmp.codpsn AND rxd.codcpc=tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Modifica
       Case 1, 2
        ' Actualizo historico de remuneraciones descuentos y aportes
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc, tmp.secuencia, tmp.codmon, tmp.importe_mn, "
        s_Sql = s_Sql & "tmp.importe_me, tmp.codcta_debmn, tmp.codcta_habmn, tmp.codcta_debme, tmp.codcta_habme, tmp.pdoano, tmp.pdomes, tmp.tipocpc, tmp.impbolecpc, tmp.codproce_pdo, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " res ON tmp.codcls=res.codcls AND tmp.codpdo=res.codpdo AND tmp.codproce=res.codproce AND tmp.codpsn=res.codpsn AND tmp.codcpc=res.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(res.codcls, res.codpdo, res.codproce, res.codpsn, res.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY codpdo, codproce, codpsn, secuencia, codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
        ' Actualizo historico de datos adicionales
        If nContador = 2 Then
          s_Sql = "INSERT INTO pldatoresultado "
          s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codpdo, tmp.codpsn, tmp.codcco, tmp.codafp, tmp.codeps, tmp.regpension, tmp.naciextrapsn, tmp.fecingreso, tmp.codubica, tmp.codsec, tmp.codcgo, tmp.fecestado, tmp.estadopsn, "
          s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN pldatoresultado dxr ON tmp.codcls=dxr.codcls AND tmp.codpdo=dxr.codpdo AND tmp.codpsn=dxr.codpsn "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(CONCAT(dxr.codcls, dxr.codpdo, dxr.codpsn), '')='' "
          s_Sql = s_Sql & "GROUP BY tmp.codcls, tmp.codpdo, tmp.codpsn "
          s_Sql = s_Sql & "ORDER BY codpdo, codpsn"
          gdl_Conexion.Execucion s_Sql, Inserta
        End If
       Case 3
        'Elimino Registro de Asistencia de personal que coincida con registtos que se desean levantar precargados en la tabla temporal
        s_Sql = "DELETE asis FROM plasistencia asis, tmpplasistencia tmp "
        s_Sql = s_Sql & "WHERE asis.codcls = tmp.codcls "
        s_Sql = s_Sql & "AND asis.codpdo=tmp.codpdo AND asis.codpsn=tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Elimina
       
        ' Insert asistencia de personal
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codpdo, tmp.codpsn, tmp.diatrabajo, tmp.diamediotm, tmp.diaparcial, tmp.dialaboral, tmp.horanormal, tmp.horamediotm, tmp.horaparcial, tmp.horatipo1, tmp.horatipo2, tmp.horatipo3, tmp.horatipo4, tmp.diafalta, "
        s_Sql = s_Sql & "tmp.tardanza, tmp.diaprepostnatal, tmp.codmdi_natal, tmp.fechaini_natal, tmp.fechafin_natal, tmp.numecitt_natal, tmp.accidente, tmp.codmdi_accid, tmp.fechaini_accid, tmp.fechafin_accid, "
        s_Sql = s_Sql & "tmp.diavacaciones, tmp.codmdi_vacac, tmp.enfermedad, tmp.codmdi_enfer, tmp.fechaini_enfer, tmp.fechafin_enfer, tmp.numecitt_enfer, tmp.licencia, tmp.codmdi_licen, tmp.fechaini_licen, tmp.fechafin_licen, "
        s_Sql = s_Sql & "tmp.diaferiado, tmp.diatradesemanal, tmp.diasuspension, tmp.dialibre, tmp.permisos, tmp.fechainivacacion, tmp.fechafinvacacion, "
        s_Sql = s_Sql & "tmp.pdovaca1, tmp.fechainivaca1, tmp.fechafinvaca1, tmp.pdovaca2, tmp.fechainivaca2, tmp.fechafinvaca2, tmp.dialiquidacion, tmp.liquidavacacion, tmp.diagratificacion, "
        s_Sql = s_Sql & "tmp.fechacese, tmp.fechainiliqvaca, tmp.fechafinliqvaca, tmp.observacion, tmp.liqnocalifica, tmp.tercerturno, tmp.opcional, "
        s_Sql = s_Sql & "tmp.diavacaventa, tmp.pdovaca3, tmp.fechainivaca3, tmp.fechafinvaca3, tmp.indvacadelanta, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " asi ON tmp.codcls=asi.codcls AND tmp.codpdo=asi.codpdo AND tmp.codpsn=asi.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(asi.codcls, asi.codpdo, asi.codpsn), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 4
        'Elimino remuneraciones descuentos exepcionales que coinciden con lo que se tiene la tabla temporal
        s_Sql = "DELETE rde FROM plremuexce rde,tmpplremuexce tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls = rde.codcls And tmp.codpdo = rde.codpdo And tmp.codpsn = rde.codpsn And tmp.codcpc = rde.codcpc"
        gdl_Conexion.Execucion s_Sql, Elimina
       
        ' Inserto remuneraciones descuentos exepcionales
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codpdo, tmp.codpsn, tmp.codcpc, tmp.codmon, tmp.imporemune, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " rde ON tmp.codcls=rde.codcls AND tmp.codpdo=rde.codpdo AND tmp.codpsn=rde.codpsn AND tmp.codcpc=rde.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(rde.codcls, rde.codpdo, rde.codpsn, rde.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn, codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
      End Select
    End If
    pgbProgreso.Value = nContador + 1
    DoEvents
  Next nContador
  ppActualiza_Procesos = True

End Function
Private Function ppActualiza_Tablas() As Boolean
  Dim sTabla As String
  Dim nContador As Integer
  Dim nRegistro As Long, nRegistros As Long

  ' Inicializo la barra de progreso
  pgbProgreso.Max = chkTabla.Count
  pgbProgreso.Value = pgbProgreso.Min
  For nContador = 0 To chkTabla.Count - 1
    ' Verifico que se haya seleccionado
    If chkTabla(nContador).Value Then
      sTabla = Choose(nContador + 1, "plbanco", "plentidadafp", "plentidadeps", "plcargo", "plprofesion", "pldocidentidad", "plconcepto", "plubicacion", "plseccion", "plctacencos", "plpersonal", "plproceso")
      
      sfmProgreso.Caption = " Actualizando Información: " & Trim(chkTabla(nContador).Caption) & " "
      Select Case nContador
       Case 0
        ' Actualizo entidad bancaria
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codbco , tmp.desbco, tmp.cuentamn, "
        s_Sql = s_Sql & "tmp.cuentame, tmp.codentidad, tmp.formato, tmp.estadobco, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " bco ON tmp.codbco=bco.codbco "
        s_Sql = s_Sql & "WHERE IFNULL(bco.codbco, '')='' "
        s_Sql = s_Sql & "ORDER BY codbco"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 1
        ' Actualizo entidad de pensión
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codafp, tmp.desafp, tmp.factor1, tmp.factor2, "
        s_Sql = s_Sql & "tmp.factor3, tmp.factor4, tmp.codbco, tmp.ctacteafp, tmp.desctacteafp, "
        s_Sql = s_Sql & "tmp.ctactefondo, tmp.desctactefondo, tmp.estadoafp, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " afp ON tmp.codafp=afp.codafp "
        s_Sql = s_Sql & "WHERE IFNULL(afp.codafp, '')='' "
        s_Sql = s_Sql & "ORDER BY codafp"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 2
        ' Actualizo entidad prestadora de salud
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codeps, tmp.deseps, tmpruceps, tmp.factoreps, tmp.estadoeps, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " eps ON tmp.codeps=eps.codeps "
        s_Sql = s_Sql & "WHERE IFNULL(eps.codeps, '')='' "
        s_Sql = s_Sql & "ORDER BY codeps"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 3
        ' Actualizo cargo de personal
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codcgo, tmp.descgo, tmp.estadocgo, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " cgo ON tmp.codcls=cgo.codcls AND tmp.codcgo=cgo.codcgo "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cgo.codcls, cgo.codcgo), '')='' "
        s_Sql = s_Sql & "ORDER BY codcgo"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 4
        ' Actualizo profesion u oficio
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codpfs, tmp.despfs, tmp.estadopfs, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " pfs ON tmp.codpfs=pfs.codpfs "
        s_Sql = s_Sql & "WHERE IFNULL(pfs.codpfs, '')='' "
        s_Sql = s_Sql & "ORDER BY codpfs"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 5
        ' Actualizo documento de identidad
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.coddci, tmp.desdci, tmp.sigladci, tmp.estadodci, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " dci ON tmp.coddci=dci.coddci "
        s_Sql = s_Sql & "WHERE IFNULL(dci.coddci, '')='' "
        s_Sql = s_Sql & "ORDER BY coddci"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 6
        ' Actualizo concepto de Cálculo
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcpc, tmp.descpc, tmp.aliascpc, tmp.tipocpc, tmp.estadocpc, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " cpc ON tmp.codcpc=cpc.codcpc "
        s_Sql = s_Sql & "WHERE IFNULL(cpc.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
        ' Actualizo concepto x planilla
        s_Sql = "INSERT INTO plconceplanilla "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codcpc, tmp.clasecpc, tmp.defaultcpc, tmp.impbolecpc, tmp.formulafun, tmp.imagenfun, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxc ON tmp.codcls=cxc.codcls AND tmp.codcpc=cxc.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcls, cxc.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 7
        ' Actualizo ubicación o localidad
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codubica, tmp.desubica, tmp.estadoubica, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " ubi ON tmp.codubica=ubi.codubica "
        s_Sql = s_Sql & "WHERE IFNULL(ubi.codubica, '')='' "
        s_Sql = s_Sql & "ORDER BY codubica"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 8
        ' Actualizo sección de la empresa
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codsec, tmp.dessec, tmp.estadosec, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " sec ON tmp.codsec=sec.codsec "
        s_Sql = s_Sql & "WHERE IFNULL(sec.codsec, '')='' "
        s_Sql = s_Sql & "ORDER BY codsec"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 9
        ' Actualizo centro de costo
        s_Sql = "INSERT INTO " & ps_DaBasCon & ".cocco "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcco, tmp.detcco, tmp.estcco, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmpcocco tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco "
        s_Sql = s_Sql & "WHERE IFNULL(cco.codcco, '')='' "
        s_Sql = s_Sql & "ORDER BY codcco"
        gdl_Conexion.Execucion s_Sql, Inserta
        ' Actualizo cuenta x concepto
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codcco, tmp.codsec, tmp.codcpc, tmp.orden, tmp.codafp, tmp.codcta_debmn, tmp.codcta_habmn, tmp.codcta_debme, tmp.codcta_habme, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " cxc ON tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codsec.cxc.codsec AND tmp.codcpc=cxc.codcpc AND tmp.orden=cxc.orden "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcls, cxc.codcco, cxc.codsec, cxc.codcpc, cxc.orden), '')='' "
        s_Sql = s_Sql & "ORDER BY codcco, codsec, codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
       Case 10
        ' Inserto Personal
        s_Sql = "INSERT INTO " & sTabla & " (codcls, codpsn, apepaterno, apematerno, nombres, fecnacimiento, ubigeonac, nacionalidad, naciextrapsn, sexopsn, refedirec, codvia, "
        s_Sql = s_Sql & "nomviadirec, numerdirec, intedirec, codzona, nomzondirec, ubigeodir, estcivilpsn, numhijo, numdepen, coddci, numdociden, "
        s_Sql = s_Sql & "numdocmil, telefono, celular, dctojudicial, pordsctojudi, fecingreso, codtpt, codcgo, cgoconfianza, codpfs, codcco, "
        s_Sql = s_Sql & "codafp, numeroafp,afpmixta, pagodolar, periodicidad, tippago, codbcopago, cuentapago, ctsdeposito, ctsdolar, codbcocts, cuentacts, codeps, siteps, regpension, fecingregpen, "
        s_Sql = s_Sql & "essvida, cobsctr, afilsindical, remintegralgrati, remintegralvaca, remintegralcts, remimprecisa, remuneta, netocpc, variacpc, imporemuneto, "
        s_Sql = s_Sql & "fecbaja, nroessalud, codubica, codsec, coddeudor, codacredor, fecestado, fotopsn, estadopsn, "
        s_Sql = s_Sql & "usrcre, fyhcre, usrmdf, fyhmdf) "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codpsn, tmp.apepaterno, tmp.apematerno, tmp.nombres, tmp.fecnacimiento, tmp.ubigeonac, tmp.nacionalidad, tmp.naciextrapsn, tmp.sexopsn, tmp.refedirec, tmp.codvia, "
        s_Sql = s_Sql & "tmp.nomviadirec, tmp.numerdirec, tmp.intedirec, tmp.codzona, tmp.nomzondirec, tmp.ubigeodir, tmp.estcivilpsn, tmp.numhijo, tmp.numdepen, tmp.coddci, tmp.numdociden, "
        s_Sql = s_Sql & "tmp.numdocmil, tmp.telefono, tmp.celular, tmp.dctojudicial, tmp.pordsctojudi, tmp.fecingreso, tmp.codtpt, tmp.codcgo, tmp.cgoconfianza, tmp.codpfs, tmp.codcco, "
        s_Sql = s_Sql & "tmp.codafp, tmp.numeroafp,tmp.afpmixta, tmp.pagodolar, tmp.periodicidad, tmp.tippago, tmp.codbcopago, tmp.cuentapago, tmp.ctsdeposito, tmp.ctsdolar, tmp.codbcocts, tmp.cuentacts, tmp.codeps, tmp.siteps, tmp.regpension, tmp.fecingregpen, "
        s_Sql = s_Sql & "tmp.essvida, tmp.cobsctr, tmp.afilsindical, tmp.remintegralgrati, tmp.remintegralvaca, tmp.remintegralcts, tmp.remimprecisa, tmp.remuneta, tmp.netocpc, tmp.variacpc, tmp.imporemuneto, "
        s_Sql = s_Sql & "tmp.fecbaja, tmp.nroessalud, tmp.codubica, tmp.codsec, tmp.coddeudor, tmp.codacredor, tmp.fecestado, tmp.fotopsn, tmp.estadopsn, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " psn ON tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(psn.codcls, psn.codpsn), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Actualizo Personal
        s_Sql = "UPDATE " & sTabla & " psn, tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "SET psn.remuneta=tmp.remuneta, psn.netocpc=tmp.netocpc, psn.variacpc=tmp.variacpc, psn.imporemuneto=tmp.imporemuneto, "
        s_Sql = s_Sql & "psn.usrmdf=tmp.usrcre, psn.fyhmdf=tmp.fyhcre "
        s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn"
        gdl_Conexion.Execucion s_Sql, Modifica
        
        ' Inserto informaicon de familiares
         s_Sql = "INSERT INTO plfamiliares (codcls,codpsn,orden, apepaterno,apematerno,nombres,fecnacimiento, sexofam, coddci,numdociden,vinculo,domicilio,codvia,nomviadom,numerdom,intedom,codzona,usrcre,fyhcre) "
         s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls,tmp.codpsn,tmp.orden, tmp.apepaterno,tmp.apematerno,tmp.nombres,tmp.fecnacimiento, tmp.sexofam, tmp.coddci,tmp.numdociden,tmp.vinculo,tmp.domicilio,tmp.codvia,tmp.nomviadom,tmp.numerdom,tmp.intedom,tmp.codzona,tmp.usrcre,tmp.fyhcre "
         s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
         s_Sql = s_Sql & "LEFT JOIN plfamiliares fam ON tmp.codcls=fam.codcls AND tmp.codpsn=fam.codpsn "
         s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
         s_Sql = s_Sql & "AND IFNULL(CONCAT(fam.codcls, fam.codpsn), '')<>'' "
         s_Sql = s_Sql & "ORDER BY tmp.codpsn"
         gdl_Conexion.Execucion s_Sql, Inserta
       Case 11
        ' Actualizo proceso de Cálculo
        s_Sql = "INSERT INTO " & sTabla & " "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codproce, tmp.desproce, tmp.estadoproce, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " prc ON tmp.codcls=prc.codcls AND tmp.codproce=prc.codproce "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(prc.codcls, prc.codproce), '')='' "
        s_Sql = s_Sql & "ORDER BY codproce"
        gdl_Conexion.Execucion s_Sql, Inserta
        ' Actualizo concepto x proceso de Cálculo
        s_Sql = "INSERT INTO plconceproceso "
        s_Sql = s_Sql & "SELECT DISTINCTROW tmp.codcls, tmp.codproce, tmp.codcpc, tmp.secuencia, tmp.formulafun, "
        s_Sql = s_Sql & "tmp.usrcre, tmp.fyhcre, tmp.usrmdf, tmp.fyhmdf "
        s_Sql = s_Sql & "FROM tmpplconceproceso tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceproceso cpr ON tmp.codcls=cpr.codcls AND tmp.codproce=cpr.codproce AND tmp.codcpc=cpr.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cpr.codcls, cpr.codproce, cpr.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY codproce, codcpc"
        gdl_Conexion.Execucion s_Sql, Inserta
      End Select
    End If
    pgbProgreso.Value = nContador + 1
    DoEvents
  Next nContador
  ppActualiza_Tablas = True
  
End Function
Private Sub ppImporta_Procesos(nPestana As Integer)
  Dim pofsoFileImp As FileSystemObject, potxtFileImp As TextStream
  Dim sTabla As String, s_Archivo As String, psRegistro As String
  Dim nRegistro As Long, nRegistros As Long
  Dim a_Tabla(), a_Archivo(), a_Columnas()
  Dim sSQLValor As String
  Dim n_Secuencia As Integer, nColumnas As Integer, nContador As Integer
  Dim s_Caracter As String, n_Elemento As Integer
  Dim a_Cabeceras(), a_Formatos(), a_Registros()
  Dim a_Cabecera(), a_Formato()
  Dim l_ExistRecord As Boolean
  Dim Nom_Tabla As String
  
  Est_CierrePeriodo = ""
  
  ' Creo objeto de archivo
  Set pofsoFileImp = CreateObject("Scripting.FileSystemObject")
  ' Importo las tablas de acuerdo a la selección
  For n_Index = 0 To chkProceso.Count - 1
    ' Verifico que se haya seleccionado
    If chkProceso(n_Index).Value Then
      ' Selecciono el nombre y columnas del archivo de texto
      a_Archivo = Choose(n_Index + 1, Array("rxd"), Array("rda"), Array("rhc"), Array("asi"), Array("rde"))
      a_Tabla = Choose(n_Index + 1, Array("plremudefa"), Array("plresultado"), Array("plresultado", "pldatoresultado"), Array("plasistencia"), Array("plremuexce"))
      a_Columnas = Choose(n_Index + 1, Array(5), Array(18), Array(29), Array(63), Array(6))
      a_Cabecera = Choose(n_Index + 1, Array("codcls", "codpsn", "codcpc", "codmon", "imporemune"), _
                   Array("codcls", "codpdo", "codproce", "codpsn", "codcpc", "secuencia", "codmon", "importe_mn", "importe_me", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme", "pdoano", "pdomes", "tipocpc", "impbolecpc", "codproce_pdo"), _
                   Array("codcls", "codpdo", "codproce", "codpsn", "codcpc", "secuencia", "codmon", "importe_mn", "importe_me", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme", "pdoano", "pdomes", "tipocpc", "impbolecpc", "codproce_pdo", "codcco", "codafp", "codeps", "regpension", "naciextrapsn", "estadopsn", "codubica", "codsec", "codcgo", "fecestado", "fecingreso"), _
                   Array("codcls", "codpdo", "codpsn", "diatrabajo", "diamediotm", "diaparcial", "dialaboral", "horanormal", "horamediotm", "horaparcial", "horatipo1", "horatipo2", "horatipo3", "horatipo4", "diafalta", "tardanza", "diaprepostnatal", "codmdi_natal", "fechaini_natal", "fechafin_natal", "numecitt_natal", "accidente", "codmdi_accid", "fechaini_accid", "fechafin_accid", "diavacaciones", "codmdi_vacac", "enfermedad", "codmdi_enfer", "fechaini_enfer", "fechafin_enfer", "numecitt_enfer", "licencia", "codmdi_licen", "fechaini_licen", "fechafin_licen", "diaferiado", "diatradesemanal", "diasuspension", "dialibre", "permisos", "fechainivacacion", "fechafinvacacion", "pdovaca1", "fechainivaca1", "fechafinvaca1", "pdovaca2", "fechainivaca2", "fechafinvaca2", "dialiquidacion", "liquidavacacion", "diagratificacion", "fechacese", "fechainiliqvaca", "fechafinliqvaca", "observacion", "liqnocalifica", "tercerturno", "opcional", "diavacaventa", "pdovaca3", "fechainivaca3", "fechafinvaca3"), _
                   Array("codcls", "codpdo", "codpsn", "codcpc", "codmon", "imporemune"))
      a_Formato = Choose(n_Index + 1, Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, _
                         TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero))
      
      ' tabla temporal en la que se cargara la data del txt
      Nom_Tabla = a_Tabla(0)
      
      ' Desactivo la opcion si no existe archivo
      chkProceso(n_Index).Value = vbUnchecked
      For n_Secuencia = 0 To UBound(a_Archivo, 1)
        ' Verifico si existe el archivo de texto y activo la opción
        s_Archivo = dlbDirectorio(nPestana).path & "\" & ps_RucEmpresa & a_Archivo(n_Secuencia) & ".sma"
        If dir$(s_Archivo, vbNormal) <> "" Then
          ' Activo la opcion si existe archivo
          chkProceso(n_Index).Value = vbChecked
          Set potxtFileImp = pofsoFileImp.OpenTextFile(s_Archivo, ForReading, False, TristateFalse)
          nRegistros = CLng(FileLen(s_Archivo))
          If nRegistros > 0 Then
            ' Redimenciono el arreglo de grabación
            nColumnas = a_Columnas(n_Secuencia)
            ReDim a_Registros(nColumnas)
            ' Inicializo la barra de progreso
            pgbProgreso.Max = nRegistros
            pgbProgreso.Value = pgbProgreso.Min
            sfmProgreso.Caption = " Importando Información: " & Trim(chkProceso(n_Index).Caption) & " - " & Right(s_Archivo, 18) & " "
            ' Elimino y creo el archivo temporal de grabacion/restauración de información
            gdl_Conexion.Execucion "DROP TABLE IF EXISTS tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1)
            s_Sql = "CREATE TEMPORARY TABLE tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1)
            s_Sql = s_Sql & " SELECT *, '999999' AS registro"
            s_Sql = s_Sql & " FROM " & a_Tabla(n_Secuencia)
            s_Sql = s_Sql & " WHERE usrcre='tmpusrma'"
            gdl_Conexion.Execucion s_Sql, Inserta
            If n_Index = 2 Then           ' Resultados historicos
              s_Sql = "ALTER TABLE tmpplresultado "
              s_Sql = s_Sql & "ADD COLUMN codcco varchar(5) Null, "
              s_Sql = s_Sql & "ADD COLUMN codafp char(2) Null, "
              s_Sql = s_Sql & "ADD COLUMN codeps char(2) Null, "
              s_Sql = s_Sql & "ADD COLUMN regpension char(1) Null, "
              s_Sql = s_Sql & "ADD COLUMN naciextrapsn char(1) Null, "
              s_Sql = s_Sql & "ADD COLUMN fecingreso date Null, "
              s_Sql = s_Sql & "ADD COLUMN codubica char(2) Null, "
              s_Sql = s_Sql & "ADD COLUMN codsec char(2) Null, "
              s_Sql = s_Sql & "ADD COLUMN codcgo char(2) Null, "
              s_Sql = s_Sql & "ADD COLUMN fecestado date Null,"
              s_Sql = s_Sql & "ADD COLUMN estadopsn varchar(1) NULL;"
              gdl_Conexion.Execucion s_Sql, Inserta
            End If
            ' Inicializo los arreglos de configuración de la tabla
            a_Cabeceras = a_Cabecera
            a_Formatos = a_Formato
            nRegistro = 0
            ' Barro todo el archivo de texto y grabo en la tabla temporal creada
            Do While Not potxtFileImp.AtEndOfStream
              nRegistro = potxtFileImp.Line
              psRegistro = potxtFileImp.ReadLine
              ' Verifico si esta dentro del rango de periodos
              l_ExistRecord = True
              ' Verifico si registro es vacio
              l_ExistRecord = (Trim(psRegistro) <> "")
              If n_Index = 2 And l_ExistRecord Then           ' Resultados historicos
                s_Caracter = Mid(psRegistro, 4, (InStr(4, psRegistro, "|") - 4))
                l_ExistRecord = (Left(psRegistro, 2) = ps_ClsPlanilla And s_Caracter >= txtPeriodo(0).Text And s_Caracter <= txtPeriodo(1).Text)
              End If
              If l_ExistRecord Then
                Registro_Texto psRegistro, nColumnas, a_Registros
                ' Genero la cadena de grabación
                s_Sql = "INSERT INTO tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1) & " ("
                sSQLValor = "VALUES("
                For nContador = 1 To nColumnas
                  n_Elemento = nContador - 1
                  s_Caracter = ", "
                  s_Sql = s_Sql & a_Cabeceras(n_Elemento) & s_Caracter
                  If a_Formatos(n_Elemento) = TipoDato.Caracter Then
                    a_Registros(nContador) = gdl_Funcion.SacaEntRetApos(a_Registros(nContador))
                    sSQLValor = sSQLValor & IIf(a_Registros(nContador) = "", "NULL", "'" & a_Registros(nContador) & "'")
                  ElseIf a_Formatos(n_Elemento) = TipoDato.FECHA Then
                    If IsDate(a_Registros(nContador)) Then
                      sSQLValor = sSQLValor & "DATE_FORMAT('" & Format(a_Registros(nContador), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "')"
                    Else
                      sSQLValor = sSQLValor & "NULL"
                    End If
                  ElseIf a_Formatos(n_Elemento) = TipoDato.Numero Then
                    a_Registros(nContador) = IIf(IsNumeric(a_Registros(nContador)), a_Registros(nContador), 0)
                    sSQLValor = sSQLValor & CDec(a_Registros(nContador))
                  End If
                  sSQLValor = sSQLValor & s_Caracter
                Next nContador
                ' Información del usuario y fecha-hora
                s_Sql = s_Sql & "registro, usrcre, fyhcre) "
                sSQLValor = sSQLValor & "'" & Format(nRegistro, "000000") & "', "
                sSQLValor = sSQLValor & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "')"
                s_Sql = s_Sql & sSQLValor
               
                ' Ejecuto la insercion del registro en la tabla
                gdl_Conexion.Execucion s_Sql, Inserta
              End If
              pgbProgreso.Value = IIf((nRegistro * 23) > nRegistros, nRegistros, (nRegistro * 23))
              DoEvents
            Loop
          End If
          potxtFileImp.Close
        End If
      Next n_Secuencia
    End If
  Next n_Index
  Set pofsoFileImp = Nothing

  If Nom_Tabla = "plasistencia" Then
    s_Sql = "SELECT codcls,codpdo FROM tmpplasistencia GROUP BY codcls,codpdo"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If (porstRecordset.BOF And porstRecordset.EOF) = False Then
      Est_CierrePeriodo = Valida_CierrePeriodo(porstRecordset!codcls, porstRecordset!codpdo, "N")
      porstRecordset.Close
    End If
  End If
  
  If Nom_Tabla = "plremuexce" Then
    s_Sql = "select codcls,codpdo from tmpplremuexce GROUP BY codcls,codpdo"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If (porstRecordset.BOF And porstRecordset.EOF) = False Then
      Est_CierrePeriodo = Valida_CierrePeriodo(porstRecordset!codcls, porstRecordset!codpdo, "N")
    End If
  End If
 
End Sub
Private Sub ppImporta_Tablas(nPestana As Integer)
  Dim pofsoFileImp As FileSystemObject, potxtFileImp As TextStream
  Dim sTabla As String, s_Archivo As String, psRegistro As String
  Dim nRegistro As Long, nRegistros As Long
  Dim a_Tabla(), a_Archivo(), a_Columnas()
  Dim sSQLValor As String
  Dim n_Secuencia As Integer, nColumnas As Integer, nContador As Integer
  Dim s_Caracter As String, n_Elemento As Integer
  Dim a_Cabeceras(), a_Formatos(), a_Registros()
  Dim a_Cabecera(), a_Formato()
  
  ' Creo objeto de archivo
  Set pofsoFileImp = CreateObject("Scripting.FileSystemObject")
  ' Importo las tablas de acuerdo a la selección
  For n_Index = 0 To chkTabla.Count - 1
    ' Verifico que se haya seleccionado
    If chkTabla(n_Index).Value Then
      ' Selecciono el nombre y columnas del archivo de texto
      a_Archivo = Choose(n_Index + 1, Array("bco"), Array("afp"), Array("eps"), Array("cgo"), Array("prf"), Array("dci"), Array("cpc", "cxp"), Array("ubi"), Array("sec"), Array("cco", "cxc"), Array("psn", "fam"), Array("prc", "cpr"))
      a_Tabla = Choose(n_Index + 1, Array("plbanco"), Array("plentidadafp"), Array("plentidadeps"), Array("plcargo"), Array("plprofesion"), Array("pldocidentidad"), Array("plconcepto", "plconceplanilla"), Array("plubicacion"), Array("plseccion"), Array(ps_DaBasCon & ".cocco", "plctacencos"), Array("plpersonal", "plfamiliares"), Array("plproceso", "plconceproceso"))
      a_Columnas = Choose(n_Index + 1, Array(7), Array(12), Array(5), Array(4), Array(3), Array(4), Array(5, 5), Array(3), Array(3), Array(3, 10), Array(69, 23), Array(4, 4))
      a_Cabecera = Choose(n_Index + 1, Array("codbco", "desbco", "cuentamn", "cuentame", "codentidad", "formato", "estadobco"), _
                   Array("codafp", "desafp", "factor1", "factor2", "factor3", "factor4", "codbco", "ctacteafp", "desctacteafp", "ctactefondo", "desctactefondo", "estadoafp"), _
                   Array("codeps", "deseps", "ruceps", "factoreps", "estadoeps"), Array("codcls", "codcgo", "descgo", "estadocgo"), _
                   Array("codpfs", "despfs", "estadopfs"), Array("coddci", "desdci", "sigladci", "estadodci"), _
                   Array(Array("codcpc", "descpc", "aliascpc", "tipocpc", "estadocpc"), Array("codcls", "codcpc", "clasecpc", "defaultcpc", "impbolecpc")), _
                   Array("codubica", "desubica", "estadoubica"), Array("codsec", "dessec", "estadosec"), _
                   Array(Array("codcco", "detcco", "estcco"), Array("codcls", "codcco", "codsec", "codcpc", "orden", "codafp", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme")), _
                   Array(Array("codcls", "codpsn", "apepaterno", "apematerno", "nombres", "fecnacimiento", "ubigeonac", "naciextrapsn", "nacionalidad", "sexopsn", "refedirec", "codvia", _
                         "nomviadirec", "numerdirec", "intedirec", "codzona", "nomzondirec", "ubigeodir", "estcivilpsn", "numhijo", "numdepen", "coddci", "numdociden", _
                         "numdocmil", "telefono", "celular", "dctojudicial", "pordsctojudi", "fecingreso", "codtpt", "codcgo", "cgoconfianza", "codpfs", "codcco", _
                         "codafp", "numeroafp", "afpmixta", "pagodolar", "periodicidad", "tippago", "codbcopago", "cuentapago", "ctsdeposito", "ctsdolar", "codbcocts", "cuentacts", "codeps", "siteps", "regpension", "fecingregpen", _
                         "essvida", "cobsctr", "afilsindical", "remintegralgrati", "remintegralvaca", "remintegralcts", "remimprecisa", "remuneta", "netocpc", "variacpc", "imporemuneto", "fecbaja", "nroessalud", "codubica", "codsec", "coddeudor", "codacredor", "fecestado", "estadopsn"), _
                         Array("codcls", "codpsn", "orden", "apepaterno", "apematerno", "nombres", "fecnacimiento", "sexofam", "coddci", "numdociden", "vinculo", "domicilio", "codvia", "nomviadom", "numerdom", "intedom", "codzona", "nomzonadom", "refedom", "ubigeodom", "incapacidad", "motivoina", "estadofam")), _
                   Array(Array("codcls", "codproce", "desproce", "estadoproce"), Array("codcls", "codproce", "codcpc", "secuencia")))
      a_Formato = Choose(n_Index + 1, Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter), Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), _
                   Array(Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)), _
                   Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), _
                   Array(Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)), _
                   Array(Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
                         TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, _
                         TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
                          TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, _
                         TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter), _
                         Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)), _
                   Array(Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter), Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero)))
      
      ' Desactivo la opcion si no existe archivo
      chkTabla(n_Index).Value = vbUnchecked
      
      For n_Secuencia = 0 To UBound(a_Archivo, 1)
        ' Verifico si existe el archivo de texto y activo la opción
        s_Archivo = dlbDirectorio(nPestana).path & "\" & ps_RucEmpresa & a_Archivo(n_Secuencia) & ".sma"
        If dir$(s_Archivo, vbNormal) <> "" Then
          ' Activo la opcion si existe archivo
          chkTabla(n_Index).Value = vbChecked
          Set potxtFileImp = pofsoFileImp.OpenTextFile(s_Archivo, ForReading, False, TristateFalse)
          nRegistros = CLng(FileLen(s_Archivo))
          If nRegistros > 0 Then
            ' Redimenciono el arreglo de grabación
            nColumnas = a_Columnas(n_Secuencia)
            ReDim a_Registros(nColumnas)
            ' Inicializo la barra de progreso
            pgbProgreso.Max = nRegistros
            pgbProgreso.Value = pgbProgreso.Min
            sfmProgreso.Caption = " Importando Información: " & Trim(chkTabla(n_Index).Caption) & " - " & Right(s_Archivo, 18) & " "
            
            ' Elimino y creo el archivo temporal de grabacion/restauración de información
            gdl_Conexion.Execucion "DROP TABLE IF EXISTS tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1)
            
            's_Sql = "CREATE TEMPORARY TABLE tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1)
            s_Sql = "CREATE TABLE tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1)
            s_Sql = s_Sql & " SELECT *, '999999' AS registro"
            s_Sql = s_Sql & " FROM " & a_Tabla(n_Secuencia)
            s_Sql = s_Sql & " WHERE usrcre='tmpusrma'"
            gdl_Conexion.Execucion s_Sql, Inserta
            
            ' Inicializo los arreglos de configuración de la tabla
            a_Cabeceras = a_Cabecera
            a_Formatos = a_Formato
            If UBound(a_Archivo, 1) > 0 Then
              a_Cabeceras = a_Cabecera(n_Secuencia)
              a_Formatos = a_Formato(n_Secuencia)
            End If
            nRegistro = 0
            ' Barro todo el archivo de texto y grabo en la tabla temporal creada
            Do While Not potxtFileImp.AtEndOfStream
              nRegistro = potxtFileImp.Line
              psRegistro = potxtFileImp.ReadLine
              Registro_Texto psRegistro, nColumnas, a_Registros
              
              ' Genero la cadena de grabación
              s_Sql = "INSERT INTO tmp" & Mid(a_Tabla(n_Secuencia), InStr(a_Tabla(n_Secuencia), ".") + 1) & " ("
              sSQLValor = "VALUES("
              For nContador = 1 To nColumnas
                n_Elemento = nContador - 1
                s_Caracter = ", "
                s_Sql = s_Sql & a_Cabeceras(n_Elemento) & s_Caracter
                If a_Formatos(n_Elemento) = TipoDato.Caracter Then
                  a_Registros(nContador) = gdl_Funcion.SacaEntRetApos(a_Registros(nContador))
                  sSQLValor = sSQLValor & IIf(a_Registros(nContador) = "", "NULL", "'" & a_Registros(nContador) & "'")
                ElseIf a_Formatos(n_Elemento) = TipoDato.FECHA Then
                  If IsDate(a_Registros(nContador)) Then
                    sSQLValor = sSQLValor & "DATE_FORMAT('" & Format(a_Registros(nContador), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "')"
                  Else
                    sSQLValor = sSQLValor & "NULL"
                  End If
                ElseIf a_Formatos(n_Elemento) = TipoDato.Numero Then
                  a_Registros(nContador) = IIf(IsNumeric(a_Registros(nContador)), a_Registros(nContador), 0)
                  sSQLValor = sSQLValor & CDec(a_Registros(nContador))
                End If
                sSQLValor = sSQLValor & s_Caracter
              Next nContador
              ' Información del usuario y fecha-hora
              s_Sql = s_Sql & "registro, usrcre, fyhcre) "
              sSQLValor = sSQLValor & "'" & Format(nRegistro, "000000") & "', "
              sSQLValor = sSQLValor & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "')"
              s_Sql = s_Sql & sSQLValor
              ' Ejecuto la insercion del registro en la tabla
              gdl_Conexion.Execucion s_Sql, Inserta
              pgbProgreso.Value = IIf((nRegistro * 23) > nRegistros, nRegistros, (nRegistro * 23))
              DoEvents
            Loop
          End If
          potxtFileImp.Close
        End If
      Next n_Secuencia
    End If
  Next n_Index
  Set pofsoFileImp = Nothing

End Sub
Private Function ppValida_Procesos(sArchivo As String) As Boolean
  Dim sTabla As String
  Dim nContador As Integer
  Dim nRegistro As Long, nRegistros As Long

  ' Inicializo la barra de progreso
  pgbProgreso.Max = chkProceso.Count
  pgbProgreso.Value = pgbProgreso.Min
  For nContador = 0 To chkProceso.Count - 1
    ' Verifico que se haya seleccionado
    If chkProceso(nContador).Value Then
      sTabla = Choose(nContador + 1, "plremudefa", "plresultado", "plresultado", "plasistencia", "plremuexce")
      sfmProgreso.Caption = " Validación de Información: " & Trim(chkProceso(nContador).Caption) & " "
      Select Case nContador
       Case 0
        ' Remuneración y descuento default existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'pk', CONCAT('Remuneración Descuento Default : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " rxd ON tmp.codcls=rxd.codcls AND tmp.codpsn=rxd.codpsn AND tmp.codcpc=rxd.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(rxd.codcls, rxd.codpsn, rxd.codcpc), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Remuneración y descuento default duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'tk', CONCAT('Remuneración Descuento Default (veces) : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc,''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.codpsn, tmp.codcpc "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Remuneración y descuento default vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'rb', CONCAT('Remuneración Descuento Default : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codpsn, tmp.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Personal no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Codigo Personal : ', IFNULL(tmp.codpsn, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(psn.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto no existente en planilla
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Concepto x Planilla : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxp ON tmp.codcls=cxp.codcls AND tmp.codcpc=cxp.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cxp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Moneda de remuneración descuelto x default no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Remuneración Descuento Default  : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codmon, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codmon, '') NOT IN('N', 'E') "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 1, 2
        ' Historico de resultado existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'pk', CONCAT('Historico Resultados (Remuneración, descuento y Aportes) : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " res ON tmp.codcls=res.codcls AND tmp.codpdo=res.codpdo AND tmp.codproce=res.codproce AND tmp.codpsn=res.codpsn AND tmp.codcpc=res.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(res.codcls, res.codpdo, res.codproce, res.codpsn, res.codcpc), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Historico de resultado duplicada en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'tk', CONCAT('Historico Resultados (Remuneración, descuento y Aportes) (veces) : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Historico de resultados vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'rb', CONCAT('Historico Resultados (Remuneración, descuento y Aportes) : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Periodo de pago no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Periodo de Pago : ', IFNULL(tmp.codpdo, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plperiodo pdo ON tmp.codcls=pdo.codcls AND tmp.codpdo=pdo.codpdo AND tmp.pdoano=pdo.anopdo "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpdo, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(pdo.codpdo, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Procesos de Cálculo no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Proceso de Cálculo : ', IFNULL(tmp.codproce, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plproceso pro ON tmp.codcls=pro.codcls AND tmp.codproce=pro.codproce "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codproce, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(pro.codproce, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce, tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Personal no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Personal de Planilla: ', IFNULL(tmp.codpsn, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(psn.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn, tmp.codpdo, tmp.codproce, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto no existente en planilla
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Concepto x Planilla : ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxp ON tmp.codcls=cxp.codcls AND tmp.codcpc=cxp.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cxp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc, tmp.codpdo, tmp.codproce, tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto no existente en proceso de Cálculo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Concepto x Proceso : ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceproceso cxp ON tmp.codcls=cxp.codcls AND tmp.codproce=cxp.codproce AND tmp.codcpc=cxp.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cxp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc, tmp.codpdo, tmp.codproce, tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Contable debe moneda nacional no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Cuenta contable Debe MN Centro Costo : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codcta_debmn, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plctacencos cxc ON tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codcpc=cxc.codcpc AND tmp.codcta_debmn=cxc.codcta_debmn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codcco, tmp.codcpc, tmp.codcta_debmn), '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcco, cxc.codcpc, cxc.codcta_debmn), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcta_debmn, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc, tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Contable haber moneda nacional no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Cuenta contable Haber MN Centro Costo : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codcta_habmn, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plctacencos cxc ON tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codcpc=cxc.codcpc AND tmp.codcta_habmn=cxc.codcta_habmn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codcco, tmp.codcpc, tmp.codcta_habmn), '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcco, cxc.codcpc, cxc.codcta_habmn), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcta_habmn, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc, tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Contable debe moneda extranjera no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Cuenta contable Debe ME Centro Costo : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codcta_debme, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plctacencos cxc ON tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codcpc=cxc.codcpc AND tmp.codcta_debme=cxc.codcta_debme "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codcco, tmp.codcpc, tmp.codcta_debme), '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcco, cxc.codcpc, cxc.codcta_debme), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcta_debme, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc, tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Contable haber moneda extranjera no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Cuenta contable Haber ME Centro Costo : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codcta_habme, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plctacencos cxc ON tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codcpc=cxc.codcpc AND tmp.codcta_habme=cxc.codcta_habme "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codcco, tmp.codcpc, tmp.codcta_habme), '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcco, cxc.codcpc, cxc.codcta_habme), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcta_habme, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc, tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Año de Cálculo no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Año de Proceso : ', IFNULL(tmp.pdoano, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.pdoano, '')<>'" & ps_Anyo & "' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Mes de Cálculo no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Mes de Proceso : ', IFNULL(tmp.pdoano, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND (IFNULL(tmp.pdomes, '')<'01' AND IFNULL(tmp.pdomes, '')>'12')"
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Tipo de concepto no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Tipo Concepto Cálculo : ', IFNULL(tmp.tipocpc, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.tipocpc, '') NOT IN('0', '1', '2') "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Imprime boleta concepto x planilla no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Impresión Concepto en Boleta : ', IFNULL(tmp.impbolecpc, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.impbolecpc, '') NOT IN('0', '1') "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Procesos de Cálculo de periodo no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Proceso de Cálculo x Periodo : ', IFNULL(tmp.codproce_pdo, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plproceso pro ON tmp.codcls=pro.codcls AND tmp.codproce_pdo=pro.codproce "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codproce_pdo, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(pro.codproce, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce_pdo, tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Moneda de historico no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Moneda de historico de Cálculo  : ', IFNULL(tmp.codmon, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codmon, '') NOT IN('N', 'E') "
        s_Sql = s_Sql & "ORDER BY tmp.codproce_pdo, tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        If nContador = 2 Then
          ' Historico de datos existente(duplicada)
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'pk', CONCAT('Historico datos de Resultados : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN pldatoresultado dxr ON tmp.codcls=dxr.codcls AND tmp.codpdo=dxr.codpdo AND tmp.codpsn=dxr.codpsn "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND IFNULL(CONCAT(dxr.codcls, dxr.codpdo, dxr.codpsn), '')<>'' "
          s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Centro de costo no existe
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Centro Costo : ', IFNULL(tmp.codcco, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.codcco, '')<>'' "
          s_Sql = s_Sql & "AND IFNULL(cco.codcco, '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Entidad de pensión - AFP no existe
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Entidad de Pensión(AFP) : ', IFNULL(tmp.codafp, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON tmp.codafp=afp.codafp "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.codafp, '')<>'' "
          s_Sql = s_Sql & "AND IFNULL(afp.codafp, '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codafp, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Entidad prestadora de salud no existe
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Entidad Prestadora de Salud : ', IFNULL(tmp.codeps, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON tmp.codeps=eps.codeps "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.codeps, '')<>'' "
          s_Sql = s_Sql & "AND IFNULL(eps.codeps, '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codeps, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Regimen de pensiones no valido
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Regimen de Pensiones : ', IFNULL(tmp.regpension, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.regpension, '') NOT IN('0', '1') "
          s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Domiciliado o No Domicialiado no valido
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Domicialiado o No Domiciliado : ', IFNULL(tmp.naciextrapsn, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.naciextrapsn, '') NOT IN('0', '1') "
          s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Fecha de ingreso de personal no valida
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Fecha Ingreso Personal : ', IFNULL(tmp.fecingreso, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.fecingreso, '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Ubicación de trabajador no existe
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Ubicación o Localidad de Trabajador : ', IFNULL(tmp.codubica, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN plubicacion ubi ON tmp.codubica=ubi.codubica "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.codubica, '')<>'' "
          s_Sql = s_Sql & "AND IFNULL(ubi.codubica, '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codubica, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Sección de empresa de trabajador no existe
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Ubicación o Localidad de Trabajador : ', IFNULL(tmp.codsec, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN plseccion sec ON tmp.codsec=sec.codsec "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.codsec, '')<>'' "
          s_Sql = s_Sql & "AND IFNULL(sec.codsec, '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codsec, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Cargo de trabajador no existe
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Cargo de Trabajador : ', IFNULL(tmp.codcgo, ''), ' - ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON tmp.codcls=cgo.codcls AND tmp.codcgo=cgo.codcgo "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.codcgo, '')<>'' "
          s_Sql = s_Sql & "AND IFNULL(CONCAT(cgo.codcls, cgo.codcgo), '')='' "
          s_Sql = s_Sql & "ORDER BY tmp.codcgo, tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
          ' Estado de trabajador no valido
          s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
          s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Estado de trabajador : ', IFNULL(tmp.estadopsn, ''), ' - ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
          s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
          s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND tmp.pdoano='" & ps_Anyo & "' "
          s_Sql = s_Sql & "AND IFNULL(tmp.estadopsn, '') NOT IN('A', 'I', 'V', 'L', 'P', 'N') "
          s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codproce, tmp.codpsn, tmp.codcpc"
          gdl_Conexion.Execucion s_Sql, Seleccion
        End If
       Case 3     ' Asistencia
        ' Personal no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Codigo Personal : ', IFNULL(tmp.codpsn, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(psn.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Periodo de pago no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Periodo de Pago : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plperiodo pdo ON tmp.codcls=pdo.codcls AND tmp.codpdo=pdo.codpdo "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpdo, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(pdo.codpdo, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 4
        ' Remuneración y descuento Exepcionales existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'pk', CONCAT('Remuneración Descuento Exepcional : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " rex ON tmp.codcls=rex.codcls AND tmp.codpdo=rex.codpdo AND tmp.codpsn=rex.codpsn AND tmp.codcpc=rex.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(rex.codcls, rex.codpdo, rex.codpsn, rex.codcpc), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Remuneración y descuento exepcional duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'tk', CONCAT('Remuneración Descuento Exepcional (veces) : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc,''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.codpdo, tmp.codpsn, tmp.codcpc "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Remuneración y descuento exepcional vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'rb', CONCAT('Remuneración Descuento Exepcional : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(tmp.codpdo, tmp.codpsn, tmp.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Periodo de pago no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Periodo de Pago : ', IFNULL(tmp.codpdo, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plperiodo pdo ON tmp.codcls=pdo.codcls AND tmp.codpdo=pdo.codpdo AND pdo.anopdo='" & ps_Anyo & "' "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpdo, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(pdo.codpdo, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Personal no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Codigo Personal : ', IFNULL(tmp.codpsn, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(psn.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto no existente en planilla
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'ne', CONCAT('Concepto x Planilla : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxp ON tmp.codcls=cxp.codcls AND tmp.codcpc=cxp.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cxp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Moneda de remuneración descuento exepcional no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Moneda de Remuneración Descuento Exepcional : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codmon, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codmon, '') NOT IN('N', 'E') "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Importe remuneración descuento exepcional no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkProceso(nContador).Caption)) & "', 'nv', CONCAT('Importe Remuneración Descuento Exepcional : ', IFNULL(tmp.codpdo, ''), ' ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.imporemune, '0.00')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.imporemune, '0.00')<=0 "
        s_Sql = s_Sql & "ORDER BY tmp.codpdo, tmp.codpsn, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
      End Select
    End If
    pgbProgreso.Value = nContador + 1
    DoEvents
  Next nContador

End Function
Private Function ppValida_Tablas(sArchivo As String) As Boolean
  Dim sTabla As String
  Dim nContador As Integer
  Dim nRegistro As Long, nRegistros As Long

  ' Inicializo la barra de progreso
  pgbProgreso.Max = chkTabla.Count
  pgbProgreso.Value = pgbProgreso.Min
  For nContador = 0 To chkTabla.Count - 1
    ' Verifico que se haya seleccionado
    If chkTabla(nContador).Value Then
      sTabla = Choose(nContador + 1, "plbanco", "plentidadafp", "plentidadeps", "plcargo", "plprofesion", "pldocidentidad", "plconcepto", "plubicacion", "plseccion", "plctacencos", "plpersonal", "plproceso")
     
      
      sfmProgreso.Caption = " Validación de Información: " & Trim(chkTabla(nContador).Caption) & " "
      Select Case nContador
       Case 0
        ' Entidad bancaria existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Entidad Banco : ', IFNULL(tmp.codbco, ''), ' ', IFNULL(tmp.desbco, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " bco ON tmp.codbco=bco.codbco "
        s_Sql = s_Sql & "WHERE IFNULL(bco.codbco, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codbco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad bancaria duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Entidad Banco (veces) : ', IFNULL(tmp.codbco, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codbco "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codbco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad bacaria vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Banco : ', IFNULL(tmp.codbco, ''), ' ', IFNULL(tmp.desbco, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codbco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codbco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Formato de transferencia no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Formato de Transferencia : ', IFNULL(tmp.formato, ''), ' - ', IFNULL(tmp.codbco, ''), ' ', IFNULL(tmp.desbco, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.formato, '') NOT IN('0', '1', '2', '3', '4') "
        s_Sql = s_Sql & "ORDER BY tmp.codbco"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 1
        ' Entidad de pensión existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Entidad de Pensión : ', IFNULL(tmp.codafp, ''), ' ', IFNULL(tmp.desafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " afp ON tmp.codafp=afp.codafp "
        s_Sql = s_Sql & "WHERE IFNULL(afp.codafp, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codafp"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad de pensión duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Entidad de Pensión (veces) : ', IFNULL(tmp.codafp, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codafp "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codafp"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad de pensión vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Entidad de Pensión : ', IFNULL(tmp.codafp, ''), ' ', IFNULL(tmp.desafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codafp, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codafp"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad bancaria no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Codigo Banco de Entidad de Pensión : ', IFNULL(tmp.codafp, ''), ' - ', IFNULL(tmp.codbco, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plbanco bco ON tmp.codbco=bco.codbco "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codbco, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(bco.codbco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codafp"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 2
        ' Entidad prestadora de salud existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Entidad de Prestadora Salud : ', IFNULL(tmp.codeps, ''), ' ', IFNULL(tmp.deseps, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " eps ON tmp.codeps=eps.codeps "
        s_Sql = s_Sql & "WHERE IFNULL(eps.codeps, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codeps"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad prestadora de salud duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Entidad Prestadora Salud (veces) : ', IFNULL(tmp.codeps, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codeps "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codeps"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad prestadora de salud vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Entidad Prestadora Salud : ', IFNULL(tmp.codeps, ''), ' ', IFNULL(tmp.deseps, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codeps, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codeps"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 3
        ' Cargo de personal existente(duplicado)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Cargo Personal : ', IFNULL(tmp.codcgo, ''), ' ', IFNULL(tmp.descgo, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " cgo ON tmp.codcls=cgo.codcls AND tmp.codcgo=cgo.codcgo "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cgo.codcls, cgo.codcgo), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codcgo"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cargo de personal duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Cargo Personal (veces) : ', IFNULL(tmp.codcgo, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codcls, tmp.codcgo "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codcgo"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cargo de personal vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Cargo Personal : ', IFNULL(tmp.codcgo, ''), ' ', IFNULL(tmp.descgo, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcgo, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcgo"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 4
        ' Profesion u oficio existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Profesión u Oficio : ', IFNULL(tmp.codpfs, ''), ' ', IFNULL(tmp.despfs, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " pfs ON tmp.codpfs=pfs.codpfs "
        s_Sql = s_Sql & "WHERE IFNULL(pfs.codpfs, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpfs"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Profesion u oficio duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Profesión u Oficio (veces) : ', IFNULL(tmp.codpfs, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codpfs "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpfs"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Profesion u oficio vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Profesión u Oficio : ', IFNULL(tmp.codpfs, ''), ' ', IFNULL(tmp.despfs, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codpfs, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpfs"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 5
        ' Documento de identidad existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Documento Identidad : ', IFNULL(tmp.coddci, ''), ' ' , IFNULL(tmp.desdci, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " dci ON tmp.coddci=dci.coddci "
        s_Sql = s_Sql & "WHERE IFNULL(dci.coddci, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.coddci"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Documento identidad duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Documento Identidad (veces) : ', IFNULL(tmp.coddci, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.coddci "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.coddci"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Documento identidad vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Documento Identidad : ', IFNULL(tmp.coddci, ''), ' ', IFNULL(tmp.desdci, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.coddci, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.coddci"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Abreviatura documento identidad vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Abreviatura Documento Identidad : ', IFNULL(tmp.coddci, ''), ' ', IFNULL(tmp.desdci, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.sigladci, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.coddci"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 6
        ' Concepto de Cálculo existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Concepto Cálculo : ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.descpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " cpc ON tmp.codcpc=cpc.codcpc "
        s_Sql = s_Sql & "WHERE IFNULL(cpc.codcpc, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto de Cálculo duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Concepto Cálculo (veces) : ', IFNULL(tmp.codcpc, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codcpc "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto de Cálculo vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Concepto Cálculo : ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.descpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Abreviatura concepto de Cálculo vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Alias Concepto Cálculo : ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.descpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.aliascpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Tipo concepto de Cálculo no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Tipo Concepto Cálculo : ', IFNULL(tmp.tipocpc, ''), ' - ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.descpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.tipocpc, '') NOT IN('0', '1', '2') "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Concepto x planilla de Cálculo existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Concepto x Clase Planilla : ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxc ON tmp.codcls=cxc.codcls AND tmp.codcpc=cxc.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcls, cxc.codcpc), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto x planilla de Cálculo duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Concepto x Clase Planilla (veces) : ', IFNULL(tmp.codcpc, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codcls, tmp.codcpc "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codcls, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto x planilla de Cálculo vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Concepto x Planilla : ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Clase concepto x planilla no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Clase Concepto x Planilla : ', IFNULL(tmp.clasecpc, ''), ' - ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.clasecpc, '') NOT IN('C', 'F') "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' caracteristica Default concepto x planilla no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Default Concepto x Planilla : ', IFNULL(tmp.defaultcpc, ''), ' - ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.defaultcpc, '') NOT IN('0', '1') "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Imprime boleta concepto x planilla no valido
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Imprime Boleta Concepto x Planilla : ', IFNULL(tmp.impbolecpc, ''), ' - ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceplanilla tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.impbolecpc, '') NOT IN('0', '1') "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 7
        ' Ubicación o localidad existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Ubicación o Localidad : ', IFNULL(tmp.codubica, ''), ' ', IFNULL(tmp.desubica, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " ubi ON tmp.codubica=ubi.codubica "
        s_Sql = s_Sql & "WHERE IFNULL(ubi.codubica, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codubica"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Ubicación o localidad duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Ubicación o Localidad (veces) : ', IFNULL(tmp.codubica, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codubica "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codubica"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Ubicación o localidad vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Ubicación o Localidad : ', IFNULL(tmp.codubica, ''), ' ', IFNULL(tmp.desubica, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codubica, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codubica"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 8
        ' Sección de empresa existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Sección de Empresa : ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.dessec, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " sec ON tmp.codsec=sec.codsec "
        s_Sql = s_Sql & "WHERE IFNULL(sec.codsec, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codsec"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Sección de empresa duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Sección de Empresa (veces) : ', IFNULL(tmp.codsec, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codsec "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codsec"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Sección de empresa vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Sección de Empresa : ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.dessec, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codsec, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codsec"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 9
        ' Centro de costo existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Centro Costo : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.detcco, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpcocco tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco "
        s_Sql = s_Sql & "WHERE IFNULL(cco.codcco, '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Centro de costo duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Centro Costo (veces) : ', IFNULL(tmp.codcco, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmpcocco tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codcco "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Centro de costo vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Centro Costo : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.detcco, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpcocco tmp "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Centro de costo no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Centro de Costo : ', IFNULL(tmp.codcco, ''), ' - ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcco, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cco.codcco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Sección de empresa no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Sección de Empresa : ', IFNULL(tmp.codsec, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plseccion sec ON tmp.codsec=sec.codsec "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codsec, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(sec.codsec, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codsec"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta x concepto existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Cuenta x Centro Costo x Concepto : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' - ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.codafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " cxc ON tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codsec=cxc.codsec AND tmp.codcpc=cxc.codcpc AND tmp.orden=cxc.orden AND tmp.codafp=cxc.codafp"
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cxc.codcls, cxc.codcco, cxc.codsec, cxc.codcpc, cxc.orden, cxc.codafp), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta x concepto duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Cuenta x Centro Costo x Concepto (veces) : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.codcls, tmp.codcco, tmp.codsec, tmp.codcpc, tmp.orden, tmp.codafp "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codcls, tmp.codcco, tmp.codsec, tmp.codcpc, tmp.orden, tmp.codafp"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta x concepto vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Cuenta x Concepto : ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' - ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE IFNULL(CONCAT(tmp.codcco, tmp.codsec, tmp.codcpc), '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad de pensión - afp no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Entidad Pensión(AFP) : ', IFNULL(tmp.codafp, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON tmp.codafp=afp.codafp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codafp, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(afp.codafp, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc, tmp.codafp"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto por planilla no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Concepto x Planilla : ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxp ON tmp.codcls=cxp.codcls AND tmp.codcpc=cxp.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cxp.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Debe de concepto mn no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Cuenta Debe MN x Concepto : ', IFNULL(tmp.codcta_debmn, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.codafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta_debmn=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcta_debmn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cta.codcta, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc, tmp.codafp, tmp.codcta_debmn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Haber de concepto mn no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Cuenta Haber MN x Concepto : ', IFNULL(tmp.codcta_habmn, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.codafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta_habmn=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcta_habmn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cta.codcta, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc, tmp.codafp, tmp.codcta_habmn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta debe de concepto me no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Cuenta Debe ME x Concepto : ', IFNULL(tmp.codcta_debme, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.codafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta_debme=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcta_debme, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cta.codcta, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc, tmp.codafp, tmp.codcta_debme"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cuenta Haber de concepto me no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Cuenta Haber ME x Concepto : ', IFNULL(tmp.codcta_habme, ''), ' - ', IFNULL(tmp.codcco, ''), ' ', IFNULL(tmp.codsec, ''), ' ', IFNULL(tmp.codcpc, ''), ' ', IFNULL(tmp.codafp, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codcta_habme=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codcta_habme, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cta.codcta, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codcco, tmp.codsec, tmp.codcpc, tmp.codafp, tmp.codcta_habme"
        gdl_Conexion.Execucion s_Sql, Seleccion
       Case 10
        ' Personal existente(duplicado)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " psn ON tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(psn.codcls, psn.codpsn), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Personal duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Personal (veces) : ', IFNULL(tmp.codpsn, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "GROUP BY tmp.codcls, tmp.codpsn "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        'Mayo 2015
        'Valida que Documento de identidad no ha sido asignado registrado en la BD
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Personal : ', IFNULL(tmp.codpsn, ''), ' - ','Id.:  ' ,tmp.numdociden ,' corresponde a: ',' ' ,per.apepaterno, ' ' ,per.apematerno), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plpersonal per ON tmp.numdociden=per.numdociden AND tmp.codcls=per.codcls "
        s_Sql = s_Sql & "GROUP BY tmp.codcls, tmp.numdociden "
        s_Sql = s_Sql & "HAVING COUNT(tmp.numdociden)>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Personal vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Apellido paterno vacio
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Apellido Paterno Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.apepaterno, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Nombres paterno vacio
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Nombres Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.nombres, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Fecha de nacimiento de personal no valida
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Fecha Nacimiento Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.fecnacimiento, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Ubicación geografica de nacimiento de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Ubicación Geografica Nacimiento Personal : ', IFNULL(tmp.ubigeonac, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON tmp.ubigeonac=ubg.codubg AND ubg.nivelubg='2' "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.ubigeonac, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(ubg.codubg, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Tipo de via domicilio de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo Via Domicilio Personal : ', IFNULL(tmp.codvia, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN pltipovia via ON tmp.codvia=via.codvia "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codvia, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(via.codvia, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Tipo de zona domicilio de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo Zona Domicilio Personal : ', IFNULL(tmp.codzona, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON tmp.codzona=zon.codzona "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codzona, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(zon.codzona, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Ubicación geografica de domicilio de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Ubicación Geografica Domicilio Personal : ', IFNULL(tmp.ubigeodir, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON tmp.ubigeodir=ubg.codubg AND ubg.nivelubg='2' "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.ubigeodir, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(ubg.codubg, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Tipo de documento de identidad personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo Documento Identidad Personal : ', IFNULL(tmp.coddci, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON tmp.coddci=dci.coddci "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.coddci, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(dci.coddci, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Fecha de ingreso de personal no valida
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'nv', CONCAT('Fecha Ingreso Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.fecingreso, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Tipo de trabajador no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo de Trabajador : ', IFNULL(tmp.codtpt, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN pltpotrabajador tpt ON tmp.codtpt=tpt.codtpt "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codtpt, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(tpt.codtpt, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Cargo de trabajador no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Cargo de Personal : ', IFNULL(tmp.codcgo, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON tmp.codcls=cgo.codcls AND tmp.codcgo=cgo.codcgo "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcgo, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cgo.codcgo, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Profesión u ofcio no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Profesión u oficio : ', IFNULL(tmp.codpfs, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plprofesion pfs ON tmp.codpfs=pfs.codpfs "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpfs, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(pfs.codpfs, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Centro de costo personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Centro costo Personal : ', IFNULL(tmp.codcco, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON tmp.codcco=cco.codcco "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcco, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cco.codcco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad de pensión - afp no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Entidad Pensión(AFP) : ', IFNULL(tmp.codafp, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON tmp.codafp=afp.codafp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codafp, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(afp.codafp, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad bancaria - pago remuneraciones no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Entidad Bancaria Pago Remuneraciones : ', IFNULL(tmp.codbcopago, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plbanco bco ON tmp.codbcopago=bco.codbco "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codbcopago, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(bco.codbco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad bancaria - CTS remuneraciones no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Entidad Bancaria Deposito CTS : ', IFNULL(tmp.codbcocts, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plbanco bco ON tmp.codbcocts=bco.codbco "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codbcocts, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(bco.codbco, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Entidad de prestadora de salud no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Entidad Prestadora Salud : ', IFNULL(tmp.codeps, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON tmp.codeps=eps.codeps "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codeps, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(eps.codeps, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto de neto a pagar  no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Concepto Neto a Pagar : ', IFNULL(tmp.netocpc, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconcepto cpc ON tmp.netocpc=cpc.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.netocpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cpc.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto de regularización no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Concepto Regularización : ', IFNULL(tmp.variacpc, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plconcepto cpc ON tmp.variacpc=cpc.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.variacpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cpc.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Codigo de ubicacion o localidad
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Ubicación o Localidad  : ', IFNULL(tmp.codubica, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plubicacion ubi ON tmp.codubica=ubi.codubica "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codubica, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(ubi.codubica, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Codigo de sección de empresas
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Sección de empresas : ', IFNULL(tmp.codsec, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN plseccion sec ON tmp.codsec=sec.codsec "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codsec, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(sec.codsec, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Codigo de deudor no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Codigo deudor : ', IFNULL(tmp.coddeudor, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocta cta ON tmp.coddeudor=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.coddeudor, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cta.codcta, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.coddeudor, tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Codigo de acreedor no existente
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Codigo acreedor : ', IFNULL(tmp.codacredor, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocta cta ON tmp.codacredor=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE IFNULL(tmp.codacredor, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cta.codcta, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codacredor, tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
       '*********Validando información Familiares del Empleado***********************
       ' Familiar de trabajador existente(duplicado)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Familiar : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN plfamiliares fam ON tmp.codcls=fam.codcls AND tmp.codpsn=fam.codpsn and tmp.numdociden=fam.numdociden "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(fam.codcls, fam.codpsn), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
         ' ***Consecutivo Familiar de trabajador existente(duplicado)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Familiar : ','Campo Orden Duplicado ->' , IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN plfamiliares fam ON tmp.codcls=fam.codcls AND tmp.codpsn=fam.codpsn and tmp.orden=fam.orden "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(fam.codcls, fam.codpsn), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Personal - Trabajador Asociado al familiar no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Codigo del Personal : ', IFNULL(tmp.codpsn, ''), ' - ', IFNULL(tmp.orden, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN plpersonal per ON tmp.codpsn=per.codpsn "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(per.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Tipo de documento de identidad personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo Documento Identidad Personal : ', IFNULL(tmp.coddci, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON tmp.coddci=dci.coddci "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.coddci, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(dci.coddci, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        'Numero de Documento de Identidad Duplicado en tabla Familiar
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Numero Documento Identidad Personal : ', IFNULL(tmp.coddci, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN plfamiliares dci ON tmp.numdociden=dci.numdociden "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.numdociden "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Tipo de zona domicilio de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo Zona Domicilio Personal : ', IFNULL(tmp.codzona, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON tmp.codzona=zon.codzona "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codzona, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(zon.codzona, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Ubicación geografica de domicilio de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Ubicación Geografica Domicilio Personal : ', IFNULL(tmp.ubigeodom, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON tmp.ubigeodom=ubg.codubg AND ubg.nivelubg='2' "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.ubigeodom, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(ubg.codubg, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion

        ' Tipo de via domicilio de personal no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Tipo Via Domicilio Personal : ', IFNULL(tmp.codvia, ''), ' - ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "LEFT JOIN pltipovia via ON tmp.codvia=via.codvia "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codvia, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(via.codvia, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
       
       ' Personal vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, ''), ' ', IFNULL(tmp.apematerno, ''), ', ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codpsn, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Apellido paterno vacio
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Apellido Paterno Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.nombres, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.apepaterno, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Nombres paterno vacio
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Nombres Personal : ', IFNULL(tmp.codpsn, ''), ' ', IFNULL(tmp.apepaterno, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplfamiliares tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.nombres, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codpsn"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        
        
        
       Case 11
        ' Procesos de Cálculo existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Proceso Cálculo : ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.desproce, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "LEFT JOIN " & sTabla & " prc ON tmp.codcls=prc.codcls AND tmp.codproce=prc.codproce "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(prc.codcls, prc.codproce), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Proceso de Cálculo duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Proceso Cálculo (veces) : ', IFNULL(tmp.codproce, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.codproce "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codproce"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Proceso de Cálculo vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Proceso Cálculo : ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.desproce, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmp" & sTabla & " tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codproce, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce"
        gdl_Conexion.Execucion s_Sql, Seleccion
        
        ' Concepto x proceso de Cálculo existente(duplicada)
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'pk', CONCAT('Concepto x Proceso Cálculo : ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codproce, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceproceso tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceproceso cpr ON tmp.codcls=cpr.codcls AND tmp.codproce=cpr.codproce AND tmp.codcpc=cpr.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(CONCAT(cpr.codcls, cpr.codproce, cpr.codcpc), '')<>'' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto x planilla de Cálculo duplicado en la importación
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'tk', CONCAT('Concepto x Proceso Cálculo (veces) : ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codcpc, ''), ' - ', COUNT(*)), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceproceso tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "GROUP BY tmp.codproce, tmp.codcpc "
        s_Sql = s_Sql & "HAVING COUNT(*)<>1 "
        s_Sql = s_Sql & "ORDER BY tmp.codproce, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto x proceso de Cálculo vacio en el archivo
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'rb', CONCAT('Codigo Concepto x Proceso Cálculo : ', IFNULL(tmp.codproce, ''), ' ', IFNULL(tmp.codcpc, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceproceso tmp "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codproce, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce"
        gdl_Conexion.Execucion s_Sql, Seleccion
        ' Concepto x proceso de Cálculo no existe
        s_Sql = "INSERT INTO tmp" & sArchivo & " (opcion, desopcion, codcaso, detalle, registro) "
        s_Sql = s_Sql & "SELECT DISTINCTROW " & Trim$(nContador) & ", '" & UCase(Trim(chkTabla(nContador).Caption)) & "', 'ne', CONCAT('Codigo Concepto x Proceso Cálculo : ', IFNULL(tmp.codcpc, ''), ' - ', IFNULL(tmp.codproce, '')), tmp.registro "
        s_Sql = s_Sql & "FROM tmpplconceproceso tmp "
        s_Sql = s_Sql & "LEFT JOIN plconceplanilla cxc ON tmp.codcls=cxc.codcls  AND tmp.codcpc=cxc.codcpc "
        s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND IFNULL(tmp.codcpc, '')<>'' "
        s_Sql = s_Sql & "AND IFNULL(cxc.codcpc, '')='' "
        s_Sql = s_Sql & "ORDER BY tmp.codproce, tmp.codcpc"
        gdl_Conexion.Execucion s_Sql, Seleccion
      End Select
    End If
    pgbProgreso.Value = nContador + 1
    DoEvents
  Next nContador

End Function
Private Sub chkProceso_Click(Index As Integer, Value As Integer)
  txtPeriodo(0).Enabled = (chkProceso(2).Value Or chkProceso(3).Value Or chkProceso(4).Value)
  txtPeriodo(1).Enabled = (chkProceso(2).Value Or chkProceso(3).Value Or chkProceso(4).Value)
  cmdHelp(0).Enabled = (chkProceso(2).Value Or chkProceso(3).Value Or chkProceso(4).Value)
  cmdHelp(1).Enabled = (chkProceso(2).Value Or chkProceso(3).Value Or chkProceso(4).Value)
End Sub

']
Private Sub cmdAction_Click(Index As Integer)
  Dim s_OldMessage As String, s_Message As String
  Dim nValidacion As Integer
  
  If Index = 1 Then Unload Me: Exit Sub
  ' Verifico que existan registros
  If chkProceso(2).Value Then
    If txtPeriodo(0).Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago Inicial", vbExclamation: txtPeriodo(0).SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago Inicial no existe; verifique", vbExclamation: txtPeriodo(0).SetFocus: Exit Sub
    If txtPeriodo(1).Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago Final", vbExclamation: txtPeriodo(1).SetFocus: Exit Sub
    If lblHelp(1) = "" Or lblHelp(1) = "???" Then Beep: MsgBox "Periodo de Pago Final no existe; verifique", vbExclamation: txtPeriodo(1).SetFocus: Exit Sub
    If Not (txtPeriodo(1).Text >= txtPeriodo(0).Text) Then Beep: MsgBox "Periodo Pago Final debe ser Mayor e Igual Periodo de Pago Inicial", vbExclamation: txtPeriodo(1).SetFocus: Exit Sub
  End If
  Beep
  If MsgBox("¿ Estás Seguro de Procesar la " & lblTitle & " ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    fMenu.panPercent.Visible = True
    Me.Height = 6770
    
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "Transferencia de Información"
    gdl_Procedure.ps_ReportName = "rptimpinforma"
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Conexion.Execucion s_Sql, Elimina
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "opcion char(2) Null, desopcion varchar(40) Null, "
    s_Sql = s_Sql & "codcaso char(2) Null, detalle varchar(200) Null, "
    s_Sql = s_Sql & "registro varchar(6) DEFAULT '0')"
    gdl_Conexion.Execucion s_Sql, Seleccion
    
    'Forzando la Creacion de la tabla temporal de plfamiliares, esto impide errores cuando al procesar la opcion padron personal
    'no haya asociado archivo txt de familiares.
    
    s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS  tmpplfamiliares "
    s_Sql = s_Sql & "( codcls char(2) Null, codpsn varchar(11) Null, orden smallint Null, apepaterno varchar(25) Null, apematerno varchar(25) Null, nombres varchar(25) Null, fecnacimiento date, "
    s_Sql = s_Sql & "sexofam char(1), coddci char(2), numdociden varchar(11), vinculo char(1), domicilio char(1), codvia char(2), nomviadom varchar(40), "
    s_Sql = s_Sql & "numerdom varchar(4), intedom varchar(4), codzona char(2), nomzonadom varchar(40), refedom varchar(50), ubigeodom varchar(6), "
    s_Sql = s_Sql & "incapacidad char(1), motivoina char(1), estadofam char(1),registro char(5) DEFAULT '0', usrcre varchar(10), fyhcre date )"
    gdl_Conexion.Execucion s_Sql, Seleccion
    
    Select Case tabRegister.Tab
     Case 0     ' Proceso de importación de información
      ' Paso 1: Realizo la importación y validación de tablas
      ppImporta_Tablas tabRegister.Tab
      fMenu.panPercent.FloodPercent = 20
      ppValida_Tablas gdl_Procedure.ps_ReportName
      fMenu.panPercent.FloodPercent = 40
      ' Paso 2: Realizo la importación y validación de procesos
      ppImporta_Procesos tabRegister.Tab
      fMenu.panPercent.FloodPercent = 60
      ppValida_Procesos gdl_Procedure.ps_ReportName
      fMenu.panPercent.FloodPercent = 80
      ' Paso 3: Verifico el resultado
      nValidacion = 0
      ' Alertas
      s_Sql = "SELECT COUNT(*) AS registros "
      s_Sql = s_Sql & " FROM tmp" & gdl_Procedure.ps_ReportName & " "
      s_Sql = s_Sql & " WHERE codcaso IN('pk')"
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If porstRecordset!registros > 0 Then
        nValidacion = 1
      End If
      ' Errores
      s_Sql = "SELECT COUNT(*) AS registros "
      s_Sql = s_Sql & " FROM tmp" & gdl_Procedure.ps_ReportName & " "
      s_Sql = s_Sql & " WHERE codcaso IN('ne', 'nv', 'rb', 'tk')"
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If porstRecordset!registros > 0 Then
        nValidacion = 2
      End If
      
      'Periodo Cerrado
       'ABRIL 2015
      If Est_CierrePeriodo = "P.CERRADO" Then
        nValidacion = 3
      End If
      'Periodo no Existe
      If Est_CierrePeriodo = "P.NO_REGISTRADO" Then
        nValidacion = 4
      End If
      
      porstRecordset.Close
      If nValidacion = 0 Then
        MsgBox "Validación de Información se completo Satisfactoriamente" & Chr$(13) & "Presione Aceptar para Iniciar la Importación de la Información", vbInformation
      ElseIf nValidacion = 1 Then
         MsgBox "Validación de Información se completo con Alertas" & Chr$(13) & "Presione Aceptar para Imprimir Reporte de Validación para que visualize las Alertas", vbExclamation
       ElseIf nValidacion = 3 Then
          MsgBox "La información a importar hace referencia a un perido cerrado" & Chr(13) & "Imposible Generar Importación de Regsitro de Asistencia" & Chr(13) & "Favor Revisar Archivo Origen", vbExclamation
       ElseIf nValidacion = 4 Then
          MsgBox "La información a importar hace referencia a un perido que no esta registrado" & Chr(13) & "Imposible Generar Importación de Regsitro de Asistencia" & Chr(13) & "Favor Revisar Archivo Origen", vbExclamation
      Else
         MsgBox "Validación de Información tiene Errores" & Chr$(13) & "Presione Aceptar para Imprimir Reporte de Validación para que pueda corregir sus Errores", vbCritical
      End If
      ' Visualizo los errores de validación
      'ABRIL 2015
      'If nValidacion <> 0 Then
      
      If nValidacion <> 0 And nValidacion <> 3 And nValidacion <> 4 Then
        ' Parametros de Impresión
        ReDim aElemento(3, 3): ReDim aElementos(2)
        ' Parametros del Reporte
        aElemento(0, 0) = ps_CodEmpresa
        aElemento(0, 1) = "": aElemento(0, 2) = ""
        ' Formulas del Reporte
        aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
        ' Parametros de campos del Reporte
        aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
        aElemento(2, 1) = "TituloReporte;" & UCase("Validación de Inportación de Información") & ";true"
        aElemento(2, 2) = "Planilla;" & ps_ClsPlanilla & " - " & Trim(ps_DesClsPlanilla) & ";true"
        ' Filtro de Formulas y Grupos del Reporte
        aElementos(0) = "": aElementos(1) = ""
        
        ' Genera la información del reporte
        s_Sql = "SELECT * "
        s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
        Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
        ' Ejecuto reporte y saco de memoria la información
        gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, 0, False, True, False, True, True, aElemento, aElementos, porstRecordset
        
        Set porstRecordset = Nothing
        ' Elimino la tabla temporal y el rango de impresion
        s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
        gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      End If
      
      ' Paso 4: Realizo la actualización de la información
      'ABRIL 2015
      'If nValidacion <> 2 Then
       
       If nValidacion <> 2 And nValidacion <> 3 And nValidacion <> 4 Then
        s_Message = IIf(nValidacion = 1, " La Validación encontro Alertas que se pueden Obviar; ", "") & "¿ Realizamos la Importación de la Información ?"
        If MsgBox(s_Message, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
          gdl_Conexion.IniciaTransaccion    'Inicia transacción
          If Not ppActualiza_Tablas() Then GoTo Error
          fMenu.panPercent.FloodPercent = 90
          If Not ppActualiza_Procesos() Then GoTo Error
          gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
        End If
      End If
      fMenu.panPercent.FloodPercent = 100
    End Select
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  Me.Height = 6290
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
    
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0, 1     ' Periodo de Pago
    tdbHelp.Columns(0).DataField = "codpdo": tdbHelp.Columns(1).DataField = "despdo"
    tdbHelp.Caption = "Periodos de Pago"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ped", "codpdo", s_Estado_Blq & ps_ClsPlanilla & ps_Anyo, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = 2300 + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = 1570
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub dlbDirectorio_Change(Index As Integer)
  flbArchivo(Index).path = dlbDirectorio(0).path
  flbArchivo(Index).Refresh
End Sub
Private Sub drbUnidad_Change(Index As Integer)
  dlbDirectorio(Index).path = drbUnidad(Index).drive
  dlbDirectorio(Index).Refresh
End Sub

Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 6290: Me.Width = 6710
  Me.Left = 1080: Me.Top = 80
  
  ' Titulo del formulario y panel
  s_TitleWindow = Me.Caption
  lblTitle = "Información del sistema"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(2, 1) = "seleccio": aElemento(2, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 1
    aElemento(n_Index, 1) = Choose(n_Index + 1, "saldacum", "cancelar")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Procesar ", "Cancelar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(1).Cancel = True
  
  drbUnidad(0).drive = ps_PathSystem
  dlbDirectorio(0).path = ps_PathSystem
  flbArchivo(0).path = dlbDirectorio(0).path
  flbArchivo(0).Pattern = ps_RucEmpresa & "*.sma"
 
 '[ Configuración el control de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codpdo", "despdo")
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
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub

Private Sub Option1_Click()
Dim Index_Chk As Byte
For Index_Chk = 0 To chkTabla.Count - 1
 chkTabla(Index_Chk).Value = vbChecked
Next

Index_Chk = 0

For Index_Chk = 0 To chkProceso.Count - 1
 chkProceso(Index_Chk).Value = vbChecked
Next
End Sub

Private Sub Option2_Click()
Dim Index_Chk As Byte

For Index_Chk = 0 To chkTabla.Count - 1
 chkTabla(Index_Chk).Value = vbUnchecked
Next

Index_Chk = 0

For Index_Chk = 0 To chkProceso.Count - 1
 chkProceso(Index_Chk).Value = vbUnchecked
Next
End Sub

Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0, 1      ' Periodo de pago
    txtPeriodo(n_IndexHelp) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo(n_IndexHelp).SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Periodo de Pago
    s_Sql = gdl_Funcion.HelpTablas("pxe", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
   Case 1     ' Entidad de banco
    s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(ColIndex).DataField, "", "")
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
Private Sub txtPeriodo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPeriodo(Index)
End Sub
Private Sub txtPeriodo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus(Index As Integer)
  lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo(Index), "PR")
End Sub
