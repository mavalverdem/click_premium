VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fExportSunat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   10290
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10290
      _Version        =   65536
      _ExtentX        =   18150
      _ExtentY        =   873
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
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   8520
         TabIndex        =   63
         Top             =   120
         Width           =   855
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
         ItemData        =   "exportsunat.frx":0000
         Left            =   720
         List            =   "exportsunat.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   120
         Width           =   2265
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   330
         Index           =   0
         Left            =   3120
         TabIndex        =   30
         Top             =   120
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   582
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "exportsunat.frx":0004
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   31
         Top             =   120
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   556
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "exportsunat.frx":0466
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   3840
         TabIndex        =   32
         Top             =   120
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "exportsunat.frx":0894
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   9440
         TabIndex        =   56
         Top             =   120
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   0
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "exportsunat.frx":0F16
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   9860
         TabIndex        =   61
         Top             =   120
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "exportsunat.frx":1598
      End
      Begin VB.Label labelruta 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4260
         TabIndex        =   57
         Top             =   135
         Width           =   855
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mes :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   135
         Width           =   450
      End
      Begin VB.Label ruta 
         BackColor       =   &H00C0C0C0&
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
         Height          =   345
         Left            =   5160
         TabIndex        =   55
         Top             =   135
         Width           =   3255
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   1365
      Left            =   0
      TabIndex        =   26
      Top             =   5880
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   2408
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
      Top             =   7320
      Width           =   10305
      _ExtentX        =   18177
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
   Begin VB.Frame frame 
      Height          =   5340
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10300
      Begin VB.CheckBox checkcontratos 
         Caption         =   "Adicionar Contratos "
         Height          =   255
         Left            =   2640
         TabIndex        =   69
         Top             =   5040
         Width           =   2535
      End
      Begin VB.CheckBox checknuevos 
         Caption         =   "Solo Trabajadores Nuevos"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   5040
         Width           =   2535
      End
      Begin VB.TextBox txtayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   405
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   62
         Text            =   "exportsunat.frx":1C1A
         Top             =   255
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.DirListBox Dir 
         Appearance      =   0  'Flat
         Height          =   2340
         Left            =   5400
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         Caption         =   "En base a Periodos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   4515
         Left            =   5280
         TabIndex        =   36
         Top             =   760
         Width           =   4950
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   66
            Top             =   4200
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   65
            Top             =   2160
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   53
            Top             =   3960
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   51
            Top             =   3480
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   49
            Top             =   3000
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   47
            Top             =   2640
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   45
            Top             =   1680
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   43
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox checkexport 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos de los Establecimientos -  Modalidad Formativa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   24
            Left            =   480
            TabIndex        =   67
            Top             =   4200
            Width           =   4440
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos del detalle de la remuneración del trabajador (rem) CTS SEMESTRAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   19
            Left            =   480
            TabIndex        =   64
            Top             =   2120
            Width           =   4245
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos del detalle de personal de terceros - SCTR (sct)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   23
            Left            =   480
            TabIndex        =   54
            Top             =   3960
            Width           =   4440
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos del detalle de comprobantes de prestadores de servicios - modalidad formativa (for)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Index           =   22
            Left            =   480
            TabIndex        =   52
            Top             =   3480
            Width           =   4245
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos del detalle de comprobantes de prestadores de servicios - cuarta categoria (4ta)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   21
            Left            =   480
            TabIndex        =   50
            Top             =   3000
            Width           =   4245
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos del detalle de la remuneración del pensionista (pen)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   495
            Index           =   20
            Left            =   480
            TabIndex        =   48
            Top             =   2600
            Width           =   4365
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos del detalle de la remuneración del trabajador (rem)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   18
            Left            =   480
            TabIndex        =   46
            Top             =   1680
            Width           =   4245
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos de los establecimientos donde labora el trabajador (tes)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   17
            Left            =   480
            TabIndex        =   44
            Top             =   1200
            Width           =   4125
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos de los días no trabajados y no subsidiados del trabajador (not)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   16
            Left            =   480
            TabIndex        =   42
            Top             =   720
            Width           =   4365
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos de los días subsidiados del trabajador (sub)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   15
            Left            =   480
            TabIndex        =   40
            Top             =   480
            Width           =   4365
         End
         Begin VB.Label lblexport 
            Caption         =   "Datos de la jornada laboral por trabajador (jor)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   38
            Top             =   240
            Width           =   4005
         End
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   5400
         TabIndex        =   25
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   23
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   22
         Top             =   4200
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   940
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox checkexport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   4560
         TabIndex        =   59
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
         Picture         =   "exportsunat.frx":1C9D
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   4920
         TabIndex        =   60
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
         Picture         =   "exportsunat.frx":231F
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de empresas a quienes destaco o desplazo personal(sdd)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   35
         Top             =   960
         Width           =   4365
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de empresas que me destacan o desplazan personal (med)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   3
         Left            =   480
         TabIndex        =   34
         Top             =   1440
         Width           =   4485
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de derechohabientes (der)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5760
         TabIndex        =   12
         Top             =   480
         Width           =   4005
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de otros empleadores (oOO)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   5760
         TabIndex        =   11
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de períodos (p00)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   10
         Top             =   4440
         Width           =   4005
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos del personal de terceros (t05)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   9
         Top             =   4200
         Width           =   4005
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos del prestador de servicios -  modalidad formativa (t04)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   8
         Top             =   3720
         Width           =   4605
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de suspensión de cuarta categoría (s00)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   7
         Top             =   3480
         Width           =   4605
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos del  prestador de servicio - cuarta categoría (t03)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   6
         Top             =   3240
         Width           =   4725
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos del pensionista (t02)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   5
         Top             =   3000
         Width           =   4005
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos del trabajador (t01)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   4005
      End
      Begin VB.Label lblexport 
         Caption         =   $"exportsunat.frx":29A1
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   4
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   4365
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de empresas a quienes destaco o desplazo personal(edd)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   4365
      End
      Begin VB.Label lblexport 
         Caption         =   "Datos de Establecimientos Propios (esp)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   4005
      End
   End
End
Attribute VB_Name = "fExportSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla
Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private as_SelRegistro(2)
Private n_Index As Integer
Dim cnn As ADODB.Connection
Dim archivos(24) As String
Dim sql As String

Private Sub RecuperaRegistros(ByVal s_Orden As String)
  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, descls, clave, estadocls,tipo "
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


Private Sub checkexport_Click(Index As Integer)
  Select Case Index
   Case 4, 11
    If checkexport(Index).Value = Checked Then
      checknuevos.Enabled = True
      checkcontratos.Enabled = False
    Else
      checknuevos.Enabled = False
      checkcontratos.Enabled = False
      checknuevos.Value = Unchecked
      checkcontratos.Value = Unchecked
    End If
   Case 5
    If checkexport(Index).Value = Checked Then
      checknuevos.Enabled = True
      checkcontratos.Enabled = True
    Else
      checknuevos.Enabled = False
      checkcontratos.Enabled = False
      checknuevos.Value = Unchecked
      checkcontratos.Value = Unchecked
    End If
  End Select
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

Private Sub Form_Load()
dir.path = "C:\"
Set cnn = New ADODB.Connection
cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
cnn.CursorLocation = adUseClient
cnn.Open

 checknuevos.Enabled = False
 checkcontratos.Enabled = False

  Dim Item As New ValueItem
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Exportación de Información a la SUNAT"
  s_TitleTable = "Clase Planilla"
  ReDim aElemento(4, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Ok", "Tipo")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcls", "descls", "estadocls", "tipo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 800, 8700, 300, 0)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbCenter, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 1)
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
  ReDim aElemento(8, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
 '[ Configuración el control de ayuda
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
  deshabilitar
End Sub

Private Sub deshabilitar()
Dim i As Integer
  For i = 0 To 23
  'Select Case i
  'Case 7, 8, 20
  '  checkexport(i).Enabled = False
  'Case Else
  'End Select
  Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Habilito la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub



Private Sub tdbRegistro_DblClick()
  ' cmdAction_Click 0
End Sub

Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub

Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub

Private Sub cmdAction_Click(Index As Integer)
Dim i As Integer
If Index = 0 Then
  For i = 0 To 24
  Select Case i
    'Case 7, 8, 20
    Case Else
      checkexport(i).Value = 1
  End Select
  Next
End If
If Index = 1 Then
  For i = 0 To 24
  Select Case i
    'Case 7, 8, 20
    Case Else
      checkexport(i).Value = 0
  End Select
  Next
End If
If Index = 2 Then
    If cmbPeriodo = "" Then Beep: MsgBox "Debe seleccionar el Periodo de Información", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    For i = 0 To 24
      If checkexport(i).Value = 1 Then
        GoTo Saltar
      End If
    Next
    MsgBox "Debe Seleccionar por lo menos un Archivo de Exportacion", vbExclamation
    Exit Sub
Saltar:
      If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
      nombredeextension
      For i = 0 To 24
        If checkexport(i).Value = 1 Then
           exportar (archivos(i))
        End If
      Next
End If
If Index = 5 Then
    dir.Refresh
    If cmdAction(5).Outline = False Then
    dir.Visible = True
    cmdAction(5).Outline = True
    Else
    dir.Visible = False
    cmdAction(5).Outline = False
    End If
End If
If Index = 6 Then
    If cmdAction(6).Outline = False Then
    txtayuda.Visible = True
    cmdAction(6).Outline = True
    Else
    txtayuda.Visible = False
    cmdAction(6).Outline = False
    End If
End If
End Sub

Private Sub nombredeextension()
Dim i As Integer
Dim extension As String
For i = 0 To 13
Select Case i
Case 0
extension = ".esp"
Case 1
extension = ".edd"
Case 2
extension = ".sdd"
Case 3
extension = ".med"
Case 4
extension = ".t00"
Case 5
extension = ".t01"
Case 6
extension = ".t02"
Case 7
extension = ".t03"
Case 8
extension = ".soo"
Case 9
extension = ".t04"
Case 10
extension = ".t05"
Case 11
extension = ".p00"
Case 12
extension = ".o00"
Case 13
extension = ".der"
End Select
archivos(i) = ps_RucEmpresa & extension
Next
For i = 14 To 24
Select Case i
Case 14
extension = ".jor"
Case 15
extension = ".sub"
Case 16
extension = ".not"
Case 17
extension = ".tes"
Case 18
extension = ".rem"
Case 19
extension = ".rem"
Case 20
extension = ".pen"
Case 21
extension = ".4ta"
Case 22
extension = ".for"
Case 23
extension = ".sct"
Case 24
extension = ".mfe"
End Select
archivos(i) = "0601" & ps_Anyo & Left(cmbPeriodo.Text, 2) & ps_RucEmpresa & extension
Next
End Sub

Private Sub exportar(arch As String)

    Dim Rst As ADODB.Recordset
    Dim R As Boolean
    Set Rst = New ADODB.Recordset
    selectsql (arch)
   
    If sql = "" Then Exit Sub
    Rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    
    If Rst.RecordCount = 0 Then
      MsgBox " No Existen datos para Generar el Archivo " & arch, vbInformation
      Exit Sub
    End If
    
    R = Recordset_a_Csv(Rst, ruta.Caption & arch)
    
    'If R Then
        MsgBox " Se generó el archivo " & arch & " correctamente en base a " & Rst.RecordCount & " Registros", vbInformation
    'End If
    If Not Rst.State = adStateOpen Then
        Rst.Close
    End If
    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If
    'If Not cnn.State = adStateOpen Then
    '    cnn.Close
    'End If
    'If Not cnn Is Nothing Then
    '    Set cnn = Nothing
    'End If
End Sub

Private Sub selectsql(arch As String)
 
 
 'PARAMETROS MES, AÑO, RUC, TIPO DE PLANILLA
 'EMPLEADO                       01
 'OBRERO                         02
 'MODALIDAD FORMATIVA            03
 'PENSIONISTAS                   04
 'PERSONAL DE TERCEROS           05
 'CUARTA CATEGORIA               06

Dim nDiasMes As Integer

nDiasMes = gdl_Funcion.NumeroDiasMes(Left(cmbPeriodo, 2), ps_Anyo)

Dim fechabaj As Date
Dim tipop As String
Dim condicion As String
Dim condicion1 As String
Dim condicion2 As String
Dim sExpresion As String

tipop = ""
condicion = "psn.codcls IN("
condicion1 = "con.codcls IN("
condicion2 = "est.codcls IN("
fechabaj = gdl_Funcion.NumeroDiasMes(Left(cmbPeriodo, 2), ps_Anyo) & "/" & Left(cmbPeriodo, 2) & "/" & ps_Anyo
For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
  tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
  If n_Index = 0 Then
    tipop = tipop & Left(tdbRegistro.Columns(3).Text, 2) & "'"
    condicion = condicion & "'" & Left(tdbRegistro.Columns(0).Text, 2) & "'"
    condicion1 = condicion1 & "'" & Left(tdbRegistro.Columns(0).Text, 2) & "'"
    condicion2 = condicion2 & "'" & Left(tdbRegistro.Columns(0).Text, 2) & "'"
  Else
    tipop = tipop & " or psn.codcls='" & Left(tdbRegistro.Columns(3).Text, 2) & "'"
    condicion = condicion & ", '" & Left(tdbRegistro.Columns(0).Text, 2) & "'"
    condicion1 = condicion1 & ", '" & Left(tdbRegistro.Columns(0).Text, 2) & "'"
    condicion2 = condicion2 & ", '" & Left(tdbRegistro.Columns(0).Text, 2) & "'"
  End If
Next n_Index
condicion = condicion & ")"
condicion1 = condicion1 & ")"
condicion2 = condicion2 & ")"

Select Case Right(arch, 3)
  Case "esp"
    sql = "select tipepr,cdgepr,desepr,indepr,case indepr when 0 then '' else tasepr end as tasepr from plestablecimientopropio where estadoepr=1"
  Case "edd"
    sql = "select codeqd,deseqd,acteqd from plempresasqdes where estadoeqd=1"
  Case "sdd"
    sql = "select ruceed,deseed,indeed,case taseed when 0 then '' else taseed end taseed from plempresaseqdes where estadoeed=1"
  Case "med"
    sql = "select codqmd,desqmd,actqmd from plempresasqmdes where estadoqmd=1"
  Case "t00"
    'EMPLEADO, OBRERO, MODALIDAD FORMATIVA, PENSIONISTA, PERSONAL DE TERCEROS, CUARTA CATEGORIA
    If (InStr(tipop, "01") + InStr(tipop, "02") + InStr(tipop, "03") + InStr(tipop, "04") + InStr(tipop, "05") + InStr(tipop, "06")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = " SELECT LEFT(dci.codsunat, 2), psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres, psn.fecnacimiento,"
      sql = sql & " CASE sexopsn WHEN 0 THEN 1 WHEN 1 THEN 2 END AS sexopsn,"
      sql = sql & " psn.nacionalidad, psn.telefono, psn.correoelect, psn.essvida, case naciextrapsn when 0 then 1 when 1 then 2 end as naciextrapsn,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.codvia end as codvia,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.nomviadirec end as nomviadirec,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.numerdirec end as numerdirec,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.intedirec end as intedirec,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.codzona end as codzona,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.nomzondirec end as nomzondirec,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.refedirec end as refedirec,"
      sql = sql & " case naciextrapsn when 0 then '' when 1 then psn.ubigeodir end as ubigeodir"
      sql = sql & " FROM plresultado res"
      sql = sql & " INNER JOIN plpersonal psn on res.codpsn=psn.codpsn"
      sql = sql & " INNER JOIN pldocidentidad dci on psn.coddci=dci.coddci"
      If checknuevos.Value = Checked Then
        sql = sql & " where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and year(psn.fecingreso)='" & ps_Anyo & "' and month(psn.fecingreso)='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      Else
        sql = sql & " where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      End If
      sql = sql & " ) Group By res.codcls , res.codpsn order by psn.apepaterno "
    End If
  Case "t01"
   'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = " select left(dci.codsunat,2), psn.numdociden,psn.codtpt,"
      sql = sql & " case chkrl when 0 then 1 when 1 then 2 end as chkrl,est.grado,"
      sql = sql & " psn.codpfs, psn.chkDIS, left(afp.codsunat,2),"
      sql = sql & " psn.fecingregpen,psn.numeroafp,"
      If checkcontratos.Value = Checked Then
        sql = sql & " cobsctr, chksctrp,'XX' as tipcon, psn.chkreg, psn.chkmax, psn.chknoc, psn.chkoiq, psn.afilsindical, psn.periodicidad,"
      Else
        sql = sql & " cobsctr, chksctrp, con.tipcon, psn.chkreg, psn.chkmax, psn.chknoc, psn.chkoiq, psn.afilsindical, psn.periodicidad,"
      End If
      sql = sql & " case psn.codeps WHEN '' then 0 WHEN '99' THEN 0 else 1 end as afieps,"
      sql = sql & " case left(eps.codsunat, 1) WHEN '9' THEN '' WHEN '' THEN '' else left(eps.codsunat,1) end as codpes,psn.siteps,psn.chkqui,"
      sql = sql & " case psn.cgoconfianza when 1 then 2 when 2 then 1 else 0 end as cgoconfianza,"
      sql = sql & " psn.tippago,psn.chkpe,psn.cmbcatocupacional,case psn.cmbtributacion when 0 then '' else psn.cmbtributacion end tribut "
      sql = sql & " from plresultado res"
      sql = sql & " left join plpersonal psn on res.codpsn=psn.codpsn"
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci"
      sql = sql & " left join plentidadafp afp on psn.codafp=afp.codafp"
      sql = sql & " left join plentidadeps eps on psn.codeps=eps.codeps"
      If checkcontratos.Value = Checked Then
      Else
        sql = sql & " left join plcontrato con on psn.codpsn=con.codpsn"
      End If
      sql = sql & " left join plestudios est on psn.codpsn=est.codpsn and est.grado = (select max(e.grado) from plestudios e where e.codpsn = psn.codpsn and  (" & condicion
      
      If checknuevos.Value = Checked Then
        sql = sql & ")) where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and year(psn.fecingreso)='" & ps_Anyo & "' and month(psn.fecingreso)='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      Else
        sql = sql & ")) where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      End If
      
      If checkcontratos.Value = Checked Then
        sql = sql & " ) Group By res.codcls , res.codpsn order by psn.apepaterno "
      Else
        sql = sql & " ) and con.estadocon= 1 Group By res.codcls , res.codpsn order by psn.apepaterno "
      End If
    End If
  Case "t02"
  'PENSIONISTA
    If (InStr(tipop, "04")) = 0 Then
      sql = ""
      Exit Sub
    Else
      Exit Sub
      'sql = "select left(dci.codsunat,2), psn.numdociden,psn.codtpt,left(afp.codsunat,2),psn.fecingregpen,psn.numeroafp,psn.siteps,psn.tippago "
      'sql = sql & " from plpersonal psn "
      'sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      'sql = sql & " left join plentidadafp afp on psn.codafp=afp.codafp "
      'sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
      'sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and (" & condicion
      'sql = sql & " ) order by psn.apepaterno"
    End If
  Case "t03"
  'CUARTA CATEGORIA
    If (InStr(tipop, "06")) = 0 Then
      sql = ""
      Exit Sub
    Else
      Exit Sub
      'sql = "select left(dci.codsunat,2), psn.numdociden,psn.numdocmil "
      'sql = sql & " from plpersonal psn "
      'sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      'sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
      'sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and (" & condicion
      'sql = sql & " ) order by psn.apepaterno"
    End If
  Case "s00"
  'CUARTA CATEGORIA
    If (InStr(tipop, "06")) = 0 Then
      sql = ""
      Exit Sub
    Else
      Exit Sub
      'sql = "select left(dci.codsunat,2), psn.numdociden,sus.numero,sus.fecha,sus.ejercicio,sus.medio "
      'sql = sql & " from plsuspensionct sus "
      'sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      'sql = sql & " left join plpersonal psn on sus.codpsn=psn.codpsn "
    End If
  Case "t04"
  'MODALIDAD FORMATIVA
    If (InStr(tipop, "03")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = " select left(dci.codsunat,2),psn.numdociden,case psn.segmedico when 0 then 1 else 2 end as seguromed,est.grado,"
      sql = sql & " psn.codpfs,psn.resfamiliar,psn.chkDIS,psn.forprofesional,psn.chknoc "
      sql = sql & " from plresultado res"
      sql = sql & " left join plpersonal psn on res.codpsn=psn.codpsn"
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci"
      sql = sql & " left join plestudios est on psn.codpsn=est.codpsn and est.grado = (select max(e.grado) from plestudios e where e.codpsn = psn.codpsn and  (" & condicion
      sql = sql & " )) where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      sql = sql & " ) Group By res.codcls , res.codpsn order by psn.apepaterno "
    End If
  Case "t05"
    sql = ""
  Case "p00"
    'EMPLEADO, OBRERO, MODALIDAD FORMATIVA, PENSIONISTA
    If (InStr(tipop, "01") + InStr(tipop, "02") + InStr(tipop, "03") + InStr(tipop, "04")) = 0 Then
      sql = ""
      Exit Sub
    Else
        
        sql = "select left(dci.codsunat,2), psn.numdociden,"
        sql = sql & " case psn.codtpt when 57 then 5 else 1 end as chkrl,"
        sql = sql & " fecingreso , fecbaja, finperiodo, concat('0',modformativa) "
        sql = sql & " from plresultado res"
        sql = sql & " inner join plpersonal psn on res.codpsn=psn.codpsn"
        sql = sql & " inner join pldocidentidad dci on psn.coddci=dci.coddci"
        If checknuevos.Value = Checked Then
            sql = sql & " where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and year(psn.fecingreso)='" & ps_Anyo & "' and month(psn.fecingreso)='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
        Else
            sql = sql & " where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
        End If
        sql = sql & " ) Group By res.codcls , res.codpsn order by psn.apepaterno "
        
    End If
  Case "o00"
    'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
        sql = "select left(dci.codsunat,2), psn.numdociden,emp.ruc,emp.razons "
        sql = sql & " from plempleadores emp "
        sql = sql & " left join plpersonal psn on psn.codpsn=emp.codpsn "
        sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
        sql = sql & " where (date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
        sql = sql & " or ((date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and (" & condicion
        sql = sql & " ) order by psn.apepaterno"
    
    End If
  Case "der"
    'EMPLEADO, OBRERO, PENSIONISTA
    If (InStr(tipop, "01") + InStr(tipop, "02") + InStr(tipop, "04")) = 0 Then
      sql = ""
      Exit Sub
    Else
    sql = "select (select left(codsunat,2) from pldocidentidad where coddci=psn.coddci), psn.numdociden,left(dci.codsunat,2),fam.numdociden,fam.apepaterno,fam.apematerno, "
    sql = sql & " fam.nombres,fam.fecnacimiento,case sexofam when 0 then 1 when 1 then 2 end as sexofam,case vinculo when 0 then 1 else vinculo end as vinculo,case fam.tipdocpaternidad when -1 then '' else fam.tipdocpaternidad end as tipdocpaternidad ,"
    sql = sql & " fam.acrepaternidad,case estadofam when 0 then 11 when 1 then 10 end as estadofam,fam.fecalta,case fam.motivoina when 0 then '' else fam.motivoina end as motivoina ,fam.fecbaja,fam.cartamed,fam.domicilio,"
    sql = sql & " fam.codvia , fam.nomviadom, fam.numerdom, fam.intedom, fam.codzona, fam.nomzonadom, fam.refedom, fam.ubigeodom"
    sql = sql & " from plfamiliares fam"
    sql = sql & " left join pldocidentidad dci on fam.coddci=dci.coddci"
    sql = sql & " left join plpersonal psn on fam.codpsn=psn.codpsn"
    sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
    sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and (" & condicion
    sql = sql & " ) order by psn.apepaterno"
    End If
  Case "jor"
    'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = "select left(dci.codsunat,2), psn.numdociden,format(truncate(sum(horanormal),1),0) as HO,"
      sql = sql & " format(((sum(horanormal)-truncate(sum(horanormal),0))*60),0) as MO,"
      sql = sql & " format(truncate((sum(horatipo1+horatipo2+horatipo3)),1),0) AS HS,"
      sql = sql & " format(((sum(horatipo1+horatipo2+horatipo3))-truncate((sum(horatipo1+horatipo2+horatipo3)),0))*60,0) AS MH from plasistencia asi"
      sql = sql & " left join plpersonal psn on asi.codpsn=psn.codpsn"
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci"
      sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
      sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and "
      sql = sql & " left(asi.codpdo,6)='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      sql = sql & " ) group by left(dci.codsunat,2), psn.numdociden order by psn.apepaterno "
    End If
  Case "sub"
    'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = ""
      For n_Index = 1 To 2
        sExpresion = Choose(n_Index, "enfer", "natal")
        sql = sql & "SELECT LEFT(dci.codsunat, 2) AS tipodoc, psn.numdociden, asi.codmdi_" & sExpresion & " AS codmdi, asi.numecitt_" & sExpresion & " AS numecitt, "
        sql = sql & "asi.fechaini_" & sExpresion & " AS fechaini, asi.fechafin_" & sExpresion & " AS fechafin "
        sql = sql & "FROM plasistencia asi "
        sql = sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
        sql = sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
        sql = sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
        sql = sql & "WHERE ((DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' AND psn.estadopsn<>'I') "
        sql = sql & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' AND psn.estadopsn='I')) "
        sql = sql & "AND " & condicion & " "
        sql = sql & "AND pdo.anopdo='" & ps_Anyo & "' "
        sql = sql & "AND pdo.mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
        sql = sql & "AND IFNULL(asi.codmdi_" & sExpresion & ", '') IN('21', '22') "
        sql = sql & "AND IFNULL(asi.fechaini_" & sExpresion & ", '')<>'' "
        sql = sql & "AND IFNULL(asi.fechafin_" & sExpresion & ", '')<>'' "
        sql = sql & "AND asi." & Choose(n_Index, "enfermedad", "diaprepostnatal") & ">=1 "
        sql = sql & IIf(n_Index = 2, "ORDER BY numdociden", "UNION ALL ")
      Next n_Index
    End If
  Case "not"
    'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = ""
      For n_Index = 1 To 3
        sExpresion = Choose(n_Index, "licen", "accid", "vacacion")
        sql = sql & "SELECT LEFT(dci.codsunat, 2) AS tipodoc, psn.numdociden, asi.codmdi_" & Left(sExpresion, 5) & " AS codmdi, "
        sql = sql & "asi.fechaini" & IIf(n_Index = 3, "", "_") & sExpresion & " AS fechaini, asi.fechafin" & IIf(n_Index = 3, "", "_") & sExpresion & " AS fechafin "
        sql = sql & "FROM plasistencia asi "
        sql = sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
        sql = sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
        sql = sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
        sql = sql & "WHERE ((DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' AND psn.estadopsn<>'I') "
        sql = sql & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' AND psn.estadopsn='I')) "
        sql = sql & "AND " & condicion & " "
        sql = sql & "AND pdo.anopdo='" & ps_Anyo & "' "
        sql = sql & "AND pdo.mespdo='" & Left(cmbPeriodo.Text, 2) & "' "
        sql = sql & "AND IFNULL(asi.codmdi_" & Left(sExpresion, 5) & ", '') NOT IN('21', '22') "
        sql = sql & "AND IFNULL(asi.fechaini" & IIf(n_Index = 3, "", "_") & sExpresion & ", '')<>'' "
        sql = sql & "AND IFNULL(asi.fechafin" & IIf(n_Index = 3, "", "_") & sExpresion & ", '')<>'' "
        sql = sql & "AND asi." & Choose(n_Index, "licencia", "accidente", "diavacaciones") & ">=1 "
        sql = sql & IIf(n_Index = 3, "ORDER BY numdociden", "UNION ALL ")
      Next n_Index
    End If
  Case "tes"
     'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
      sql = "select left(dci.codsunat,2), psn.numdociden,case length(ruc) when 3 then '" & ps_RucEmpresa & "' when 2 then '" & ps_RucEmpresa & "' when 1 then '" & ps_RucEmpresa & "' else ruc end as ruc,codest,case tasa when 0 then '' else tasa end as tasa "
      sql = sql & " from plestalaboral "
      sql = sql & " left join plpersonal psn on plestalaboral.codpsn=psn.codpsn "
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
      sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and (" & condicion
      sql = sql & " ) and mes='" & Left(cmbPeriodo, 2) & "' and ano='" & ps_Anyo & "'  order by psn.apepaterno"
    End If
  Case "rem"
    'EMPLEADO, OBRERO
    If (InStr(tipop, "01") + InStr(tipop, "02")) = 0 Then
      sql = ""
      Exit Sub
    Else
      If checkexport(18).Value = 1 Then
        sql = "SELECT LEFT(dci.codsunat, 2), psn.numdociden, con.codsunat, "
        sql = sql & "CONVERT(ROUND(SUM(res.importe_mn), 2), char(15)) AS importe_mn1, "
        sql = sql & "CONVERT(ROUND(SUM(res.importe_mn), 2), char(15)) AS importe_mn2 "
        sql = sql & "FROM plresultado res "
        sql = sql & "LEFT JOIN plconceplanilla con on res.codcpc=con.codcpc and res.codcls=con.codcls "
        sql = sql & "LEFT JOIN plpersonal psn on res.codpsn=psn.codpsn and res.codcls=psn.codcls  "
        sql = sql & "LEFT JOIN pldocidentidad dci on psn.coddci=dci.coddci "
        sql = sql & "WHERE res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
        sql = sql & ") AND con.codsunat not in('0100','0200','0300','0400','0500','0600','0604','0607','0610','0612','0700','0800','0802','0804','0806','0808') "
        sql = sql & "GROUP BY left(dci.codsunat,2),psn.numdociden , con.codsunat , res.codcls  "
        sql = sql & "ORDER BY psn.numdociden"
      End If
      If checkexport(19).Value = 1 Then
        sql = "SELECT LEFT(dci.codsunat, 2), psn.numdociden, con.codsunat, "
        sql = sql & "CONVERT(ROUND(SUM(res.importe_mn), 2), char(17)) AS importe_mn1, "
        sql = sql & "CONVERT(ROUND(SUM(res.importe_mn), 2), char(17)) AS importe_mn2 "
        sql = sql & "from plctsresultado res "
        sql = sql & "left join plctsperiodosub per on res.codcls=per.codcls and res.pdocts=per.pdocts and res.subcts=per.subcts  "
        sql = sql & "left join plconceplanilla con on res.codcpc=con.codcpc and res.codcls=con.codcls "
        sql = sql & "left join plpersonal psn on res.codpsn=psn.codpsn and res.codcls=psn.codcls  "
        sql = sql & "left join pldocidentidad dci on psn.coddci=dci.coddci "
        sql = sql & "where year(per.fechaven)='" & ps_Anyo & "' and month(per.fechaven)='" & Left(cmbPeriodo.Text, 2) & "' and con.codsunat='0904' "
        sql = sql & "group by left(dci.codsunat,2),psn.numdociden , con.codsunat , res.codcls  "
        sql = sql & " order by psn.numdociden"
      End If
    End If
  Case "pen"
    'PENSIONISTA
     If (InStr(tipop, "04")) = 0 Then
      sql = ""
      Exit Sub
     Else
      sql = "select left(dci.codsunat,2), psn.numdociden,con.codsunat,case sum(res.importe_mn) when 0 then '0.00' else sum(res.importe_mn) end as importe_mn1 ,case sum(res.importe_mn) when 0 then '0.00' else sum(res.importe_mn) end as importe_mn2 "
      sql = sql & " from plresultado res "
      sql = sql & " left join plconceplanilla con on res.codcpc=con.codcpc "
      sql = sql & " left join plpersonal psn on res.codpsn=psn.codpsn "
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
      sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and "
      sql = sql & " res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      sql = sql & " ) and con.codsunat not in('0100','0200','0300','0400','0500','0600','0604','0607','0610','0612','0700','0800','0802','0804','0806','0808') "
      sql = sql & " group by left(dci.codsunat,2),psn.numdociden , con.codsunat "
      sql = sql & " order by psn.apepaterno"
    End If
  Case "4ta"
    'CUARTA CATEGORIA
     If (InStr(tipop, "06")) = 0 Then
      sql = ""
      Exit Sub
     Else
      sql = "select left(dci.codsunat,2), psn.numdociden,com.tipo,com.serie,com.numero,com.monto,com.fecemision,com.fecpago,com.retencion "
      sql = sql & " from plcomprobantect com "
      sql = sql & " left join plpersonal psn on com.codpsn=psn.codpsn "
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
    End If
  Case "for"
    'MODALIDAD FORMATIVA
     If (InStr(tipop, "03")) = 0 Then
      sql = ""
      Exit Sub
     Else
      sql = "select left(dci.codsunat,2), psn.numdociden,case sum(res.importe_mn) when 0 then '0.00' else sum(res.importe_mn) end as importe_mn1 "
       sql = sql & " from plresultado res "
      sql = sql & " left join plconceplanilla con on res.codcpc=con.codcpc and res.codcls=con.codcls "
      sql = sql & " left join plpersonal psn on res.codpsn=psn.codpsn and res.codcls=psn.codcls  "
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      sql = sql & " where res.pdoano='" & ps_Anyo & "' and res.pdomes='" & Left(cmbPeriodo.Text, 2) & "' and (" & condicion
      sql = sql & " ) and con.codsunat not in('0100','0200','0300','0400','0500','0600','0604','0607','0610','0612','0700','0800','0802','0804','0806','0808') "
      sql = sql & " group by left(dci.codsunat,2),psn.numdociden , con.codsunat , res.codcls  "
      sql = sql & " order by psn.numdociden"
    End If
  Case "sct"
    sql = ""
  Case "mfe"
    'MODALIDAD FORMATIVA
     If (InStr(tipop, "03")) = 0 Then
      sql = ""
      Exit Sub
     Else
      sql = "select left(dci.codsunat,2), psn.numdociden,case length(ruc) when 3 then '" & ps_RucEmpresa & "' when 2 then '" & ps_RucEmpresa & "' when 1 then '" & ps_RucEmpresa & "' else ruc end as ruc,codest "
      sql = sql & " from plestalaboral "
      sql = sql & " left join plpersonal psn on plestalaboral.codpsn=psn.codpsn "
      sql = sql & " left join pldocidentidad dci on psn.coddci=dci.coddci "
      sql = sql & " where ((date_format(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn<>'I') "
      sql = sql & " or (date_format(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cmbPeriodo.Text, 2) & "' and psn.estadopsn='I')) and (" & condicion
      sql = sql & " ) and mes='" & Left(cmbPeriodo, 2) & "' and ano='" & ps_Anyo & "'  order by psn.apepaterno"
     End If
  Case Else
    sql = ""
End Select
End Sub

Function Recordset_a_Csv(rs As Recordset, path As String) As Boolean
    On Error GoTo Err_function
    Dim Columna
    Dim Fila As Integer
    ' Crea el archivo
    Open path For Output As #1
    ' Se mueve al primer registro
    rs.MoveFirst
    ' recorre todo el recordset
    For Fila = 0 To rs.RecordCount - 1
        ' nombre del campo
        Print #1, Trim(rs.Fields(0));
        ' recorre todos los campos
        For Columna = 1 To rs.Fields.Count - 1
          ' imprime la fila actual en el fichero
          Print #1, "|" & Trim(rs.Fields(Columna));
        Next
            ' escribe una línea en blanco
        Print ""
            ' salto de carro
        Print #1, "|" & Chr(13) & Chr(10);
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

