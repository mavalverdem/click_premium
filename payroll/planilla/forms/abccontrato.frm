VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcContrato 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4965
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "abccontrato.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   8910
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2070
      TabIndex        =   38
      Top             =   4275
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
   Begin TabDlg.SSTab tabRegister 
      Height          =   3765
      Left            =   75
      TabIndex        =   37
      Top             =   600
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   6641
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
      TabPicture(0)   =   "abccontrato.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblHelp(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDato(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdHelp(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtpFechas(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpFechas(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmCuadro(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmCuadro(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtObservacion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbPeriodo(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbPeriodo(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbPeriodo(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtTipo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txtTipo 
         Height          =   300
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1305
         Width           =   375
      End
      Begin VB.ComboBox cmbPeriodo 
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
         Index           =   2
         ItemData        =   "abccontrato.frx":0028
         Left            =   4905
         List            =   "abccontrato.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   645
      End
      Begin VB.ComboBox cmbPeriodo 
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
         Index           =   1
         ItemData        =   "abccontrato.frx":002C
         Left            =   3285
         List            =   "abccontrato.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   1590
      End
      Begin VB.ComboBox cmbPeriodo 
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
         Index           =   0
         ItemData        =   "abccontrato.frx":0030
         Left            =   2415
         List            =   "abccontrato.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   840
      End
      Begin VB.TextBox txtObservacion 
         Height          =   300
         Left            =   1380
         TabIndex        =   13
         Top             =   1650
         Width           =   6450
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
         ForeColor       =   &H00FF8080&
         Height          =   300
         Left            =   1380
         TabIndex        =   1
         Top             =   270
         Width           =   990
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1005
         Index           =   1
         Left            =   6255
         TabIndex        =   14
         Top             =   270
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1773
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
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   15
            Top             =   345
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "&Activo"
            ForeColor       =   12582912
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
         Begin Threed.SSOption optEstado 
            Height          =   180
            Index           =   1
            Left            =   195
            TabIndex        =   16
            Top             =   690
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "&Vencido"
            ForeColor       =   12582912
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
      Begin Threed.SSFrame frmCuadro 
         Height          =   1305
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   2010
         Width           =   7650
         _Version        =   65536
         _ExtentX        =   13494
         _ExtentY        =   2302
         _StockProps     =   14
         Caption         =   " Documento "
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
         Begin VB.TextBox txtPlantilla 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   870
            Width           =   6440
         End
         Begin VB.TextBox txtUbicacion 
            Height          =   300
            Left            =   180
            TabIndex        =   19
            Top             =   480
            Width           =   6440
         End
         Begin Threed.SSCommand cmdArchivo 
            Height          =   360
            Index           =   0
            Left            =   6705
            TabIndex        =   20
            Top             =   435
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   12632256
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
            AutoSize        =   2
            Picture         =   "abccontrato.frx":0034
         End
         Begin Threed.SSCommand cmdArchivo 
            Height          =   360
            Index           =   1
            Left            =   7140
            TabIndex        =   21
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
            AutoSize        =   2
            Picture         =   "abccontrato.frx":0050
         End
         Begin VB.Label lblDato 
            Caption         =   "Archivo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   18
            Top             =   270
            Width           =   825
         End
      End
      Begin MSComCtl2.DTPicker dtpFechas 
         Height          =   300
         Index           =   0
         Left            =   1380
         TabIndex        =   6
         Top             =   615
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   137625601
         CurrentDate     =   37515
      End
      Begin MSComCtl2.DTPicker dtpFechas 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   8
         Top             =   960
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   137625601
         CurrentDate     =   37515
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   1860
         TabIndex        =   39
         Top             =   1305
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
         Caption         =   "Observación :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1695
         Width           =   1200
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
         Left            =   2220
         TabIndex        =   11
         Top             =   1350
         Width           =   195
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Inicio :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Termino :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   1200
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Contrato :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Número :"
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
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   1200
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8910
      _Version        =   65536
      _ExtentX        =   15716
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
         Left            =   8085
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
         Picture         =   "abccontrato.frx":006C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   7695
         TabIndex        =   25
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
         Picture         =   "abccontrato.frx":0088
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
         Left            =   675
         TabIndex        =   26
         Top             =   120
         Width           =   6420
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   27
      Top             =   4455
      Width           =   8910
      _Version        =   65536
      _ExtentX        =   15716
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
         Left            =   5595
         TabIndex        =   28
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
         Picture         =   "abccontrato.frx":00A4
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   5205
         TabIndex        =   29
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
         Picture         =   "abccontrato.frx":00C0
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   3495
         TabIndex        =   30
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
         Picture         =   "abccontrato.frx":00DC
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   3105
         TabIndex        =   31
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
         Picture         =   "abccontrato.frx":00F8
      End
      Begin MSAdodcLib.Adodc dcaHelp 
         Height          =   330
         Left            =   0
         Top             =   120
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
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   3765
      Index           =   0
      Left            =   8145
      TabIndex        =   32
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6641
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
         TabIndex        =   33
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
         Left            =   120
         TabIndex        =   34
         Tag             =   "0"
         Top             =   600
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
         Picture         =   "abccontrato.frx":0114
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   35
         Tag             =   "0"
         Top             =   1230
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
         Picture         =   "abccontrato.frx":0130
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   36
         Tag             =   "0"
         Top             =   1830
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
         Picture         =   "abccontrato.frx":014C
      End
   End
End
Attribute VB_Name = "fAbcContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Cuonter As Integer, s_ParCodigo As String     ' Indice para bucle, y parametro de codigo
Private l_Registro As Boolean                           ' Si generación archivo de contrato
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private n_Index As Integer
Private Type wrdMergeDocumento                          ' Informacion generar combinacion contrato
  Sexo As String
  Sigla As String
  nombre As String
  Edad As Integer
  EstadoCivil As String
  TipoDococumento As String
  Documento As String
  Domicilio As String
  Distrito As String
  Cargo As String
  FechaInicio As String
  FechaFinal As String
  FechaFirma As String
  PlazoMeses As Single
  Moneda As String
  Remunera As Double
  Letras As String
End Type
'[
Private Sub Bloqueo_Contrato(o_AppWord As Object)

'  Documents.Open FileName:="0877274120050621.doc", ConfirmConversions:=True, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
'   PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:=""
'  Windows(1).Activate
'  With ActiveDocument.MailMerge
'    .Destination = wdSendToNewDocument
'    .SuppressBlankLines = True
'    With .DataSource
'      .FirstRecord = wdDefaultFirstRecord
'      .LastRecord = wdDefaultLastRecord
'    End With
'    .Execute Pause:=False
'  End With
'  Windows("0877274120050621.doc").Activate
'  ActiveWindow.Close
'  ActiveDocument.SaveAs FileName:="0877274120050621.doc", FileFormat:=wdFormatDocument, LockComments:=True, Password:="sysmavm", _
'   AddToRecentFiles:=False, WritePassword:="", ReadOnlyRecommended:=True, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
'  ActiveWindow.Close

End Sub
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Cuonter = 0 To 3: cmdMove(n_Cuonter).Visible = (Me.Tag = s_MdoData_Vis): Next n_Cuonter
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdArchivo(0).Enabled = (Me.Tag = s_MdoData_Ins)
  cmdArchivo(1).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)

End Sub
Private Sub MailMerge_Documento(o_AppWord As Object)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim s_Caracter As String, s_FileTexto As String
  Dim psColumna As String, psRegistro As String
  Dim o_MergeInformacion As wrdMergeDocumento
  
  ' Obtengo las remuneraciones
  s_Sql = "SELECT cpc.descpc, rxd.imporemune, "
  s_Sql = s_Sql & "IF(rxd.codmon='" & s_Codmon_mn & "', '" & s_Codmon_mn_Txt & "', '" & s_Codmon_me_Txt & "') AS moneda "
  s_Sql = s_Sql & "FROM plremudefa rxd "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON rxd.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE rxd.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND rxd.codpsn='" & Trim(fAbcPersonal.txtCodigo) & "' "
  s_Sql = s_Sql & "AND cpc.tipocpc='" & s_Estado_Ina & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  s_Caracter = "|"
  psColumna = "": psRegistro = ""
  o_MergeInformacion.Remunera = 0: n_Cuonter = 1
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    While Not porstRecordset.EOF
      o_MergeInformacion.Moneda = porstRecordset("moneda")
      psColumna = psColumna & "moneda" & Format(n_Cuonter, "00") & s_Caracter & "remunera" & Format(n_Cuonter, "00") & s_Caracter & "letras" & Format(n_Cuonter, "00") & s_Caracter
      psRegistro = psRegistro & porstRecordset("moneda") & s_Caracter & FormatNumber(CDec(porstRecordset("imporemune")), 2) & s_Caracter & gdl_Funcion.NumeroEnLetras(CDec(porstRecordset("imporemune"))) & o_MergeInformacion.Moneda & s_Caracter
      o_MergeInformacion.Remunera = o_MergeInformacion.Remunera + CDec(porstRecordset("imporemune"))
      n_Cuonter = n_Cuonter + 1
      porstRecordset.MoveNext
    Wend
    porstRecordset.Close
  End If

  '[ Genero el archivo de texto con los datos
  s_FileTexto = ps_PathSystem & "contrato.txt"
  Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
  Set potxtFileExp = pofsoFileExp.CreateTextFile(s_FileTexto, True)
  o_MergeInformacion.Sexo = fAbcPersonal.cmbSexo
  o_MergeInformacion.Sigla = IIf(fAbcPersonal.cmbSexo.ListIndex = 0, "Don", "Doña")
  o_MergeInformacion.nombre = UCase(Trim(fAbcPersonal.txtNombres(2)) & " " & Trim(fAbcPersonal.txtNombres(0)) & " " & Trim(fAbcPersonal.txtNombres(1)))
  o_MergeInformacion.TipoDococumento = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, fAbcPersonal.txtTipoDocu.Text, "DS")
  o_MergeInformacion.Documento = txtCodigo.Text
  o_MergeInformacion.Domicilio = Trim(fAbcPersonal.lblHelp(1)) & " " & Trim(fAbcPersonal.txtNombreVia) & " " & Trim(fAbcPersonal.txtnumero(0)) & IIf(Trim(fAbcPersonal.txtnumero(1)) <> "", " - ", "") & Trim(fAbcPersonal.txtnumero(1))
  o_MergeInformacion.Domicilio = o_MergeInformacion.Domicilio & " " & Trim(fAbcPersonal.lblHelp(2)) & " " & Trim(fAbcPersonal.txtNombreZona)
  o_MergeInformacion.Distrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, fAbcPersonal.txtUbigeo(1).Text, "UB")
  o_MergeInformacion.Cargo = Trim(fAbcPersonal.lblHelp(4))
  o_MergeInformacion.FechaInicio = Format(dtpFechas(0), "dd") & " de " & Format(dtpFechas(0), "mmmm") & " del " & Format(Year(dtpFechas(0)), "#,###")
  o_MergeInformacion.FechaFinal = Format(dtpFechas(1), "dd") & " de " & Format(dtpFechas(1), "mmmm") & " del " & Format(Year(dtpFechas(1)), "#,###")
  o_MergeInformacion.FechaFirma = cmbPeriodo(2) & " día(s) del mes de " & Trim(Mid(cmbPeriodo(1), 6)) & " del " & LCase(gdl_Funcion.NumeroEnLetras(CDec(cmbPeriodo(0)), False))
  o_MergeInformacion.Moneda = IIf(fAbcPersonal.chkPagoDolar.Value, s_Codmon_me_Txt, s_Codmon_mn_Txt)
  n_Cuonter = InStr(fAbcPersonal.cmbEstadoCivil, "(")
  o_MergeInformacion.EstadoCivil = IIf(n_Cuonter <> 0, IIf(fAbcPersonal.cmbSexo.ListIndex = 0, "o", "a"), "")
  n_Cuonter = IIf(n_Cuonter <> 0, n_Cuonter - 2, Len(fAbcPersonal.cmbEstadoCivil))
  o_MergeInformacion.EstadoCivil = Left(fAbcPersonal.cmbEstadoCivil, n_Cuonter) & o_MergeInformacion.EstadoCivil
  o_MergeInformacion.Letras = UCase(gdl_Funcion.NumeroEnLetras(o_MergeInformacion.Remunera) & " " & IIf(fAbcPersonal.chkPagoDolar.Value, s_Codmon_me_Nom, s_Codmon_mn_Nom))
  
  ' Nombre de los campos
  For n_Cuonter = 1 To 16
    psColumna = psColumna & Choose(n_Cuonter, "sexo", "sigla", "nombre", "edad", "estadocivil", "tipodoc", "documento", "domicilio", "distrito", "cargo", "fechaini", "fechafin", "fechafirma", "moneda", "remunera", "letras") & IIf(n_Cuonter = 16, "", s_Caracter)
    psRegistro = psRegistro & Choose(n_Cuonter, o_MergeInformacion.Sexo, o_MergeInformacion.Sigla, o_MergeInformacion.nombre, Left(fAbcPersonal.lblEdad, 2), o_MergeInformacion.EstadoCivil, o_MergeInformacion.TipoDococumento, o_MergeInformacion.Documento, o_MergeInformacion.Domicilio, o_MergeInformacion.Distrito, o_MergeInformacion.Cargo, o_MergeInformacion.FechaInicio, o_MergeInformacion.FechaFinal, o_MergeInformacion.FechaFirma, o_MergeInformacion.Moneda, FormatNumber(o_MergeInformacion.Remunera, 2), o_MergeInformacion.Letras) & IIf(n_Cuonter = 16, "", s_Caracter)
  Next n_Cuonter
  potxtFileExp.WriteLine psColumna
  potxtFileExp.WriteLine psRegistro
  potxtFileExp.Close
  Set potxtFileExp = Nothing
  Set pofsoFileExp = Nothing
  ']

  '[ Realizo la combinación de los datos
  o_AppWord.ActiveDocument.MailMerge.MainDocumentType = nFormLetters
  o_AppWord.ActiveDocument.MailMerge.OpenDataSource Name:=s_FileTexto, ConfirmConversions:=True, _
  ReadOnly:=False, LinkToSource:=True, AddToRecentFiles:=False, PasswordDocument:="", _
  PasswordTemplate:="", WritePasswordDocument:="", WritePasswordTemplate:="", _
  Revert:=False, Format:=nOpenFormatAuto, Connection:="", SQLStatement:="SELECT * FROM contrato", _
  SQLStatement1:="", SubType:=nMergeSubTypeOther
  
  With o_AppWord.ActiveDocument.MailMerge
    .Destination = nNewBlankDocument
    .SuppressBlankLines = True
    .DataSource.FirstRecord = .DataSource.ActiveRecord
    .DataSource.LastRecord = .DataSource.ActiveRecord
    .Execute Pause:=False
  End With
  o_AppWord.Documents(2).Close nNewBlankDocument, , False
  ']
  If gdl_Funcion.ExisteArchivo(s_FileTexto) Then
    Kill s_FileTexto
  End If

End Sub
Private Sub MarkerMerge_Documento(o_AppWord As Object)
  Dim s_Caracter As String
  Dim psColumna As String, psRegistro As String
  Dim o_MergeInformacion As wrdMergeDocumento
  
  ' Obtengo las remuneraciones
  s_Sql = "SELECT cpc.descpc, rxd.imporemune, "
  s_Sql = s_Sql & "IF(rxd.codmon='" & s_Codmon_mn & "', '" & s_Codmon_mn_Txt & "', '" & s_Codmon_me_Txt & "') AS moneda "
  s_Sql = s_Sql & "FROM plremudefa rxd "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON rxd.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE rxd.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND rxd.codpsn='" & Trim(fAbcPersonal.txtCodigo) & "' "
  s_Sql = s_Sql & "AND cpc.tipocpc='" & s_Estado_Ina & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  s_Caracter = "|"
  psColumna = "": psRegistro = ""
  o_MergeInformacion.Remunera = 0: n_Cuonter = 1
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    While Not porstRecordset.EOF
      o_MergeInformacion.Moneda = porstRecordset("moneda")
      psColumna = psColumna & "moneda" & Format(n_Cuonter, "00") & s_Caracter & "remunera" & Format(n_Cuonter, "00") & s_Caracter & "letras" & Format(n_Cuonter, "00") & s_Caracter
      psRegistro = psRegistro & porstRecordset("moneda") & s_Caracter & FormatNumber(CDec(porstRecordset("imporemune")), 2) & s_Caracter & gdl_Funcion.NumeroEnLetras(CDec(porstRecordset("imporemune"))) & o_MergeInformacion.Moneda & s_Caracter
      o_MergeInformacion.Remunera = o_MergeInformacion.Remunera + CDec(porstRecordset("imporemune"))
      n_Cuonter = n_Cuonter + 1
      porstRecordset.MoveNext
    Wend
    porstRecordset.Close
  End If

  '[ Genero el archivo de texto con los datos
  o_MergeInformacion.Sexo = fAbcPersonal.cmbSexo
  o_MergeInformacion.Sigla = IIf(fAbcPersonal.cmbSexo.ListIndex = 0, "Don", "Doña")
  o_MergeInformacion.nombre = UCase(Trim(fAbcPersonal.txtNombres(2)) & " " & Trim(fAbcPersonal.txtNombres(0)) & " " & Trim(fAbcPersonal.txtNombres(1)))
  o_MergeInformacion.TipoDococumento = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, fAbcPersonal.txtTipoDocu.Text, "DS")
  o_MergeInformacion.Documento = txtCodigo.Text
  o_MergeInformacion.Domicilio = Trim(fAbcPersonal.lblHelp(1)) & " " & Trim(fAbcPersonal.txtNombreVia) & " " & Trim(fAbcPersonal.txtnumero(0)) & IIf(Trim(fAbcPersonal.txtnumero(1)) <> "", " - ", "") & Trim(fAbcPersonal.txtnumero(1))
  o_MergeInformacion.Domicilio = o_MergeInformacion.Domicilio & " " & Trim(fAbcPersonal.lblHelp(2)) & " " & Trim(fAbcPersonal.txtNombreZona)
  o_MergeInformacion.Distrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, fAbcPersonal.txtUbigeo(1).Text, "UB")
  o_MergeInformacion.Cargo = Trim(fAbcPersonal.lblHelp(4))
  o_MergeInformacion.FechaInicio = Format(dtpFechas(0), "dd") & " de " & Format(dtpFechas(0), "mmmm") & " del " & Format(Year(dtpFechas(0)), "#,###")
  o_MergeInformacion.FechaFinal = Format(dtpFechas(1), "dd") & " de " & Format(dtpFechas(1), "mmmm") & " del " & Format(Year(dtpFechas(1)), "#,###")
  o_MergeInformacion.FechaFirma = cmbPeriodo(2) & " día(s) del mes de " & LCase(Trim(Mid(cmbPeriodo(1), 6))) & " del " & LCase(gdl_Funcion.NumeroEnLetras(CDec(cmbPeriodo(0)), False))
  o_MergeInformacion.PlazoMeses = gdl_Funcion.DiferenciaFechas(dtpFechas(0), dtpFechas(1), "M")
  o_MergeInformacion.Moneda = IIf(fAbcPersonal.chkPagoDolar.Value, s_Codmon_me_Txt, s_Codmon_mn_Txt)
  n_Cuonter = InStr(fAbcPersonal.cmbEstadoCivil, "(")
  o_MergeInformacion.EstadoCivil = IIf(n_Cuonter <> 0, IIf(fAbcPersonal.cmbSexo.ListIndex = 0, "o", "a"), "")
  n_Cuonter = IIf(n_Cuonter <> 0, n_Cuonter - 2, Len(fAbcPersonal.cmbEstadoCivil))
  o_MergeInformacion.EstadoCivil = Left(fAbcPersonal.cmbEstadoCivil, n_Cuonter) & o_MergeInformacion.EstadoCivil
  o_MergeInformacion.Letras = UCase(gdl_Funcion.NumeroEnLetras(o_MergeInformacion.Remunera) & " " & IIf(fAbcPersonal.chkPagoDolar.Value, s_Codmon_me_Nom, s_Codmon_mn_Nom))
  ']

  '[ Realizo la combinacion de marcadores
  With o_AppWord.ActiveDocument
    .Bookmarks("Nombre_Empleado1").Range.Text = o_MergeInformacion.nombre
    .Bookmarks("Tipo_documento1").Range.Text = o_MergeInformacion.TipoDococumento
    .Bookmarks("Num_Documento1").Range.Text = o_MergeInformacion.Documento
    .Bookmarks("Direccion").Range.Text = o_MergeInformacion.Domicilio
    .Bookmarks("Distrito").Range.Text = o_MergeInformacion.Distrito
    .Bookmarks("Cargo_Empleado").Range.Text = o_MergeInformacion.Cargo
    .Bookmarks("Plazo_Contrato").Range.Text = o_MergeInformacion.PlazoMeses & " meses"
    .Bookmarks("Fecha_Inicio1").Range.Text = o_MergeInformacion.FechaInicio
    .Bookmarks("Fecha_Final").Range.Text = o_MergeInformacion.FechaFinal
    .Bookmarks("Fecha_Inicio2").Range.Text = o_MergeInformacion.FechaInicio
    .Bookmarks("Moneda").Range.Text = o_MergeInformacion.Moneda
    .Bookmarks("Remuneracion").Range.Text = Format(o_MergeInformacion.Remunera, s_FormatoNum_0)
    .Bookmarks("Remu_Letras").Range.Text = o_MergeInformacion.Letras
    .Bookmarks("Fecha_Firma").Range.Text = o_MergeInformacion.FechaFirma
    .Bookmarks("Nombre_Empleado2").Range.Text = o_MergeInformacion.nombre
    .Bookmarks("Tipo_documento2").Range.Text = o_MergeInformacion.TipoDococumento
    .Bookmarks("Num_Documento2").Range.Text = o_MergeInformacion.Documento
  End With
  ']

End Sub
Sub ShowScreen()
  Dim sArchivoDot As String, sArchivoDoc As String
  
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  sArchivoDot = "": sArchivoDoc = ""
  If Me.Tag = s_MdoData_Ins Then
    '[ Obtengo las ubicaciones de los archivos
    s_Sql = "SELECT cfg.contrato_dot, cfg.contrato_doc "
    s_Sql = s_Sql & "FROM plcfgempresa cfg "
    s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.EOF) Then
      sArchivoDot = gdl_Funcion.aTexto(porstRecordset!contrato_dot)
      sArchivoDoc = gdl_Funcion.aTexto(porstRecordset!contrato_doc)
      porstRecordset.Close
    End If
    ']
    gdl_Procedure.EditText "PK", txtCodigo, fAbcPersonal.txtDocumento(0).Text, s_MdoData_Upd, True, 8
    gdl_Procedure.EditCombo "PK", cmbPeriodo(0), 20, Me.Tag, False
    n_Cuonter = Month(Date) - 1
    gdl_Procedure.EditCombo "PK", cmbPeriodo(1), n_Cuonter, Me.Tag, False
    n_Cuonter = Day(Date) - 1
    gdl_Procedure.EditCombo "PK", cmbPeriodo(2), n_Cuonter, Me.Tag, False
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txttipo, "", Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtObservacion, "", Me.Tag, False, 50
    gdl_Procedure.EditText "PK", txtUbicacion, sArchivoDoc, Me.Tag, False, 100
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
    l_Registro = False
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fAbcPersonal.tdbContrato.Columns(0).Text, Me.Tag, True, 8
    n_Cuonter = Mid(fAbcPersonal.tdbContrato.Columns(0).Text, InStr(fAbcPersonal.tdbContrato.Columns(0).Text, "-") + 1, 4)
    n_Cuonter = IIf((Val(ps_Anyo) - n_Cuonter) >= 0, (20 - Abs(Val(ps_Anyo) - n_Cuonter)), (20 + Abs(Val(ps_Anyo) - n_Cuonter)))
    gdl_Procedure.EditCombo "PK", cmbPeriodo(0), n_Cuonter, Me.Tag, True
    n_Cuonter = Mid(fAbcPersonal.tdbContrato.Columns(0).Text, InStr(fAbcPersonal.tdbContrato.Columns(0).Text, "-") + 5, 2)
    gdl_Procedure.EditCombo "PK", cmbPeriodo(1), (n_Cuonter - 1), Me.Tag, True
    n_Cuonter = Right(fAbcPersonal.tdbContrato.Columns(0).Text, 2)
    gdl_Procedure.EditCombo "PK", cmbPeriodo(2), (n_Cuonter - 1), Me.Tag, True
    gdl_Procedure.EditDTPicker "PK", dtpFechas(0), fAbcPersonal.tdbContrato.Columns(1).Text, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "PK", dtpFechas(1), fAbcPersonal.tdbContrato.Columns(2).Text, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txttipo, gdl_Funcion.aTexto(fAbcPersonal.tdbContrato.Columns(6).Text), Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtObservacion, gdl_Funcion.aTexto(fAbcPersonal.tdbContrato.Columns(3).Text), Me.Tag, False, 50
    gdl_Procedure.EditText "PK", txtUbicacion, gdl_Funcion.aTexto(fAbcPersonal.tdbContrato.Columns(5).Text), Me.Tag, False, 100
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fAbcPersonal.tdbContrato.Columns(4).Value = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fAbcPersonal.tdbContrato.Columns(4).Value = s_Estado_Ina), Me.Tag, True
    l_Registro = True
  End If
  txtPlantilla.Text = sArchivoDot
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txttipo, "TC")

End Sub
']
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   cmbPeriodo(0).SetFocus
  End If
  If Index <> 1 Then Exit Sub
    
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim$(txtObservacion) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpsn", "numdocumen", "ano", "mes", "dia")
    a_Valores = Array(ps_ClsPlanilla, Trim(fAbcPersonal.txtCodigo), Trim(txtCodigo.Text), Trim(cmbPeriodo(0)), Left(Trim(cmbPeriodo(1)), 2), Trim(cmbPeriodo(2)))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plcontrato", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Elimino el archivo fisico si tiene problemas de permiso da error
    If gdl_Funcion.ExisteArchivo(txtUbicacion.Text) Then
      'Kill Trim(txtUbicacion.Text)
    End If
    ' Refresco el Ado control y la grilla
    fAbcPersonal.RecuperarContrato
    Me.Tag = s_MdoData_Vis
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

Private Sub cmdArchivo_Click(Index As Integer)
  Dim o_AppWord As Object, o_DocWord As Object
  Dim s_Archivo As String
  
'  Dim o_AppWord As Word.Application
'  Dim o_DocWord As Word.Document

  If Index = 0 Then     ' Selección de plantilla
    s_Archivo = Trim(IIf(txtPlantilla.Text = "", ps_PathSystem, txtPlantilla.Text))
    fMenu.cdlDialogo.CancelError = False
    fMenu.cdlDialogo.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist
    fMenu.cdlDialogo.Filter = "Archivos de Word(*.doc; *.docx; *.dot)|*.doc;*.docx;*.dot|Archivos de Texto(*.txt)|*.txt"
    fMenu.cdlDialogo.FilterIndex = 1
    fMenu.cdlDialogo.DialogTitle = "Seleccionar Plantilla"
    fMenu.cdlDialogo.InitDir = s_Archivo
    fMenu.cdlDialogo.FileName = ""
    fMenu.cdlDialogo.ShowOpen
    ' Capturo archivo seleccionado
    If fMenu.cdlDialogo.FileName <> "" Then
        txtPlantilla.Text = fMenu.cdlDialogo.FileName
    End If
  Else
    ' Realizo las validaciones de los campos a actualizar
    If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
    If cmbPeriodo(0) = "" Then Beep: MsgBox "Debe Ingresar Ejercicio " & lblTitle, vbExclamation: cmbPeriodo(0).SetFocus: Exit Sub
    If cmbPeriodo(1) = "" Then Beep: MsgBox "Debe Ingresar Mes " & lblTitle, vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If cmbPeriodo(2) = "" Then Beep: MsgBox "Debe Ingresar Dia de Firma " & lblTitle, vbExclamation: cmbPeriodo(2).SetFocus: Exit Sub
    If txtObservacion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtObservacion.SetFocus: Exit Sub
    If txtUbicacion = "" Then Beep: MsgBox "Debe Ingresar la la ubicación del archivo " & lblTitle, vbExclamation: txtUbicacion.SetFocus: Exit Sub
    If dir(txtUbicacion, vbDirectory) = "" Then Beep: MsgBox "Ubicación del archivo " & lblTitle & "; No valida Verificar", vbExclamation: txtUbicacion.SetFocus: Exit Sub
    If (txtPlantilla.Text = "" And Me.Tag = s_MdoData_Ins) Then Beep: MsgBox "Debe Seleccionar la plantilla de " & lblTitle, vbExclamation: cmdArchivo(0).SetFocus: Exit Sub

    Set o_AppWord = CreateObject("Word.Application.15")
    s_Archivo = Trim(IIf(Me.Tag = s_MdoData_Ins, txtPlantilla.Text, txtUbicacion.Text))
    If gdl_Funcion.ExisteArchivo(s_Archivo) Then
      If Me.Tag = s_MdoData_Ins Then
        If MsgBox("¿ Estás Seguro de Generar Archivo fisico de contrato ? ", vbQuestion + vbYesNo) = vbYes Then
          o_AppWord.Visible = True
          o_AppWord.Documents.Open (s_Archivo)
          o_AppWord.Visible = False
          MarkerMerge_Documento o_AppWord
          Set o_DocWord = o_AppWord.ActiveDocument
          ' Genero la grabación del documento
          s_Archivo = Trim(txtUbicacion.Text)
          s_Archivo = s_Archivo & txtCodigo.Text & "-" & cmbPeriodo(0).Text & Left(cmbPeriodo(1).Text, 2) & cmbPeriodo(2).Text & ".docx"
          o_DocWord.SaveAs2 FileName:=s_Archivo, FileFormat:=wdFormatDocumentDefault, LockComments:=False, Password:=txtCodigo.Text, _
              AddToRecentFiles:=False, WritePassword:="", ReadOnlyRecommended:=True, EmbedTrueTypeFonts:=False, _
              SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
          
          s_Archivo = Replace(s_Archivo, ".docx", ".pdf")
          o_DocWord.SaveAs2 FileName:=s_Archivo, FileFormat:=wdFormatPDF, LockComments:=False, Password:=txtCodigo.Text, _
              AddToRecentFiles:=False, WritePassword:="", ReadOnlyRecommended:=True, EmbedTrueTypeFonts:=False, _
              SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
          s_Archivo = Replace(s_Archivo, ".pdf", ".docx")
          
          txtUbicacion.Text = s_Archivo
          o_DocWord.Activate
          ' Realizo la actualización de información
          l_Registro = True
          cmdUpdate_Click
        End If
      Else
        Set o_DocWord = o_AppWord.Documents.Open(s_Archivo, , True, False, txtCodigo.Text, , , , , , , True)
      End If
    End If
    o_AppWord.Visible = True
    Set o_DocWord = Nothing
    Set o_AppWord = Nothing
  End If

End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fAbcPersonal.tdbContrato.MoveFirst
   Case 1: If Not fAbcPersonal.tdbContrato.BOF Then fAbcPersonal.tdbContrato.MovePrevious
           If fAbcPersonal.tdbContrato.BOF Then fAbcPersonal.tdbContrato.MoveFirst
   Case 2: If Not fAbcPersonal.tdbContrato.EOF Then fAbcPersonal.tdbContrato.MoveNext
           If fAbcPersonal.tdbContrato.EOF Then fAbcPersonal.tdbContrato.MoveLast
   Case 3: fAbcPersonal.tdbContrato.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1
  Dim s_Archivo As String

  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If cmbPeriodo(0) = "" Then Beep: MsgBox "Debe Ingresar Ejercicio " & lblTitle, vbExclamation: cmbPeriodo(0).SetFocus: Exit Sub
  If cmbPeriodo(1) = "" Then Beep: MsgBox "Debe Ingresar Mes " & lblTitle, vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
  If cmbPeriodo(2) = "" Then Beep: MsgBox "Debe Ingresar Dia de Firma " & lblTitle, vbExclamation: cmbPeriodo(2).SetFocus: Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha de termino debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  If txttipo.Text = "" Then Beep: MsgBox "Debe Ingresar el Tipo de Contrato " & lblTitle, vbExclamation: txttipo.SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Tipo de Contrato no valido; Verificar", vbExclamation: txttipo.SetFocus: Exit Sub
  If txtObservacion.Text = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtObservacion.SetFocus: Exit Sub
  If txtUbicacion.Text <> "" And Not gdl_Funcion.ExisteArchivo(Trim(txtUbicacion.Text)) Then Beep: MsgBox "Debe Ingresar la ubicación del archivo " & lblTitle & "; No valida Verificar", vbExclamation: txtUbicacion.SetFocus: Exit Sub
  If txtUbicacion.Text <> "" And Not l_Registro Then Beep: MsgBox "Debe generar el archivo fisico " & lblTitle, vbExclamation: cmdArchivo(1).SetFocus: Exit Sub
  
  s_Estado = IIf(optEstado(0).Value, s_Estado_Act, s_Estado_Ina)
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Archivo = Replace(Trim(txtUbicacion.Text), "\", "\\")
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpsn", "numdocumen", "ano", "mes", "dia", "fechaini", "fechafin", "observacion", "archivo", "estadocon", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"), "tipcon")
  a_Valores = Array(ps_ClsPlanilla, Trim(fAbcPersonal.txtCodigo), Trim(txtCodigo.Text), Trim(cmbPeriodo(0)), Left(Trim(cmbPeriodo(1)), 2), Trim(cmbPeriodo(2)), Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), Trim(txtObservacion.Text), s_Archivo, s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0), Trim(txttipo.Text))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpsn", "numdocumen", "ano", "mes", "dia")
 
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plcontrato", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plcontrato", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  fAbcPersonal.RecuperarContrato
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    Me.Tag = s_MdoData_Upd
    gdl_Procedure.EditText "PK", txtCodigo, txtCodigo.Text, Me.Tag, True, 8
    gdl_Procedure.EditCombo "PK", cmbPeriodo(0), cmbPeriodo(0).ListIndex, Me.Tag, True
    gdl_Procedure.EditCombo "PK", cmbPeriodo(1), cmbPeriodo(1).ListIndex, Me.Tag, True
    gdl_Procedure.EditCombo "PK", cmbPeriodo(2), cmbPeriodo(2).ListIndex, Me.Tag, True
    gdl_Procedure.EditDTPicker "PK", dtpFechas(0), Trim(dtpFechas(0)), Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "PK", dtpFechas(1), Trim(dtpFechas(1)), Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txttipo, txttipo.Text, Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtObservacion, txtObservacion.Text, Me.Tag, False, 50
    gdl_Procedure.EditText "PK", txtUbicacion, txtUbicacion.Text, Me.Tag, False, 100
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
    
    txtPlantilla.Text = ""
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
  fMenu.cmbejercicio.Enabled = False
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 5340: Me.Width = 9000
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización de Contratos de Trabajo"
  lblTitle = "Contrato de Trabajo"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  n_IndexHelp = -1

  ' Obtengo el modo de operación del registro
  Me.Tag = fAbcPersonal.tdbContrato.Tag
  
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
  
  ' Configuro los Controles de adicionales
  cmdArchivo(0).Outline = False: cmdArchivo(1).Outline = False
  gdl_Procedure.LoadGrafics cmdArchivo(0), "buscarch", "Seleccionar Plantilla " & lblTitle
  gdl_Procedure.LoadGrafics cmdArchivo(1), "wordlnk", "Examinar Archivo " & lblTitle
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fAbcPersonal.tdbContrato.EOF Or fAbcPersonal.tdbContrato.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fAbcPersonal.tdbContrato.Columns(0).Text
  
  ' Configuro los listados, datos adicionales
  For n_Cuonter = (Val(ps_Anyo) - 20) To (Val(ps_Anyo) + 20): cmbPeriodo(0).AddItem Format(n_Cuonter, "0000"): Next n_Cuonter
  For n_Cuonter = 1 To 12: cmbPeriodo(1).AddItem Choose(n_Cuonter, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Cuonter
  For n_Cuonter = 1 To 31: cmbPeriodo(2).AddItem Format(n_Cuonter, "00"): Next n_Cuonter
  
  ' Carga los datos en el formulario
  ShowScreen
    
   '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codtco", "destco")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Tipo de Contrato de Trabajo", 2
  ' Asigno el control de datos  ala grilla
  tdbHelp.DataSource = dcaHelp
  
  ' Recupero la información
  s_Sql = gdl_Funcion.HelpTablas("tco", tdbHelp.Columns(0).DataField, "", "")
  
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaHelp, tdbHelp, s_Sql, adCmdText, adLockReadOnly
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  fMenu.cmbejercicio.Enabled = True
  ' Habilito/desabilito botones inciales
  fAbcPersonal.cmdActionCon(0).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActionCon(1).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActionCon(2).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  Me.Tag = s_MdoData_Vis
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

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
      cmbPeriodo(0).SetFocus
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txtObservacion_GotFocus()
  gdl_Procedure.MarcaGet txtObservacion
End Sub
Private Sub cmdHelp_Click(Index As Integer)

  If Not cmdHelp(Index).Enabled Then Exit Sub
  s_SqlHelp = ""
  If n_IndexHelp = Index Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + frmCuadro(1).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub

Private Sub txttipo_GotFocus()
  gdl_Procedure.MarcaGet txttipo
End Sub
Private Sub txttipo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txttipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtUbicacion.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txttipo_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txttipo, "TC")
End Sub
Private Sub tdbHelp_DblClick()

  If dcaHelp.Recordset.RecordCount = 0 Or (dcaHelp.Recordset.EOF And dcaHelp.Recordset.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  txttipo.Text = tdbHelp.Columns(0).Value
  lblHelp(0) = tdbHelp.Columns(1).Value
  txtUbicacion.SetFocus

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  s_Sql = gdl_Funcion.HelpTablas("tco", tdbHelp.Columns(ColIndex).DataField, "", "")
  dcaHelp.RecordSource = s_Sql
  dcaHelp.Refresh

End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaHelp, tdbHelp, ""
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim n_Columna As Integer
  
  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    If Chr$(KeyAscii) = "*" Then
      s_SqlHelp = ""
    Else
      s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    End If
    n_Columna = tdbHelp.Col
    s_Sql = gdl_Funcion.HelpTablas("tco", tdbHelp.Columns(n_Columna).DataField, "", s_SqlHelp)
    dcaHelp.RecordSource = s_Sql
    dcaHelp.Refresh
    tdbHelp.Col = n_Columna
  Else
      s_SqlHelp = ""
  End If

End Sub
Private Sub tdbHelp_LostFocus()
  tdbHelp.Visible = False
End Sub
