VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcDistribuCenCos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7200
   Icon            =   "abcdiscencos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   7200
   Begin TabDlg.SSTab tabRegister 
      Height          =   2610
      Left            =   75
      TabIndex        =   20
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4604
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
      TabPicture(0)   =   "abcdiscencos.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblHelp(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdHelp(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCenCosto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPorcentaje"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtPorcentaje 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   285
         TabIndex        =   4
         Top             =   1395
         Width           =   1050
      End
      Begin VB.TextBox txtCenCosto 
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
         Left            =   285
         TabIndex        =   1
         Top             =   630
         Width           =   1100
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   1455
         TabIndex        =   19
         Top             =   630
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
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
         Index           =   0
         Left            =   1815
         TabIndex        =   2
         Top             =   675
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Porcentaje :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   3
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Caption         =   "Centro Costo :"
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
         Left            =   285
         TabIndex        =   0
         Top             =   315
         Width           =   1335
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Left            =   6450
         TabIndex        =   6
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
         Picture         =   "abcdiscencos.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   7
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
         Picture         =   "abcdiscencos.frx":0044
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
         TabIndex        =   8
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   3270
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Left            =   4695
         TabIndex        =   10
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
         Picture         =   "abcdiscencos.frx":0060
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   11
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
         Picture         =   "abcdiscencos.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   12
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
         Picture         =   "abcdiscencos.frx":0098
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   13
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
         Picture         =   "abcdiscencos.frx":00B4
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   2610
      Index           =   0
      Left            =   6435
      TabIndex        =   14
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   4604
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
         TabIndex        =   15
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edici�n"
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
         TabIndex        =   16
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
         Picture         =   "abcdiscencos.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   17
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
         Picture         =   "abcdiscencos.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   18
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
         Picture         =   "abcdiscencos.frx":0108
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   21
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
Attribute VB_Name = "fAbcDistribuCenCos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificaci�n de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_CentroCosto As String                         ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acci�n
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creaci�n", IIf(Me.Tag = s_MdoData_Del, " - Eliminaci�n", IIf(Me.Tag = s_MdoData_Upd, " - Actualizaci�n", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins)

End Sub
Sub ShowScreen()
    
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCenCosto, "", Me.Tag, False, fDistribucionCenCosto.dcaRegistro.Recordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtPorcentaje, FormatNumber(0, 2), Me.Tag, False, 7, vbRightJustify
  Else
    gdl_Procedure.EditText "PK", txtCenCosto, fDistribucionCenCosto.dcaRegistro.Recordset!codcco, Me.Tag, True, fDistribucionCenCosto.dcaRegistro.Recordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtPorcentaje, FormatNumber(fDistribucionCenCosto.dcaRegistro.Recordset!porcentaje, 2), Me.Tag, False, 18, vbRightJustify
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, ps_CodEmpresa, txtCenCosto.Text, "CC")

End Sub
']
Private Sub cmdAction_Click(Index As Integer)
  Dim s_Periodo As String, s_Personal As String
  
  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCenCosto.SetFocus
  ElseIf Index = 2 Then
    txtPorcentaje.SetFocus
  End If
  If Index <> 1 Then Exit Sub
    
  Beep
  If MsgBox("� Est�s Seguro de Eliminar la " & lblTitle & " '" & lblHelp(0) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_CentroCosto = Trim$(txtCenCosto.Text)
    s_Periodo = o_SelDisCenCosto.txtPeriodo.Text
    s_Personal = o_SelDisCenCosto.dcaRegistro.Recordset!codpsn
    
    '[ Inicio la conexi�n a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpdo", "codpsn", "codcco")
    a_Valores = Array(ps_ClsPlanilla, s_Periodo, s_Personal, Trim(txtCenCosto.Text))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    
    gdl_Conexion.IniciaTransaccion    'Inicia transacci�n
    ' Elimino el registro
    If Not Records_Del("plcencospro", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacci�n
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fDistribucionCenCosto.dcaRegistro, fDistribucionCenCosto.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fDistribucionCenCosto.dcaRegistro.Recordset.EOF And fDistribucionCenCosto.dcaRegistro.Recordset.BOF) Or fDistribucionCenCosto.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fDistribucionCenCosto.dcaRegistro.Recordset.Find ("codcco >= '" & s_CentroCosto & "'")
      If fDistribucionCenCosto.dcaRegistro.Recordset.EOF Then fDistribucionCenCosto.dcaRegistro.Recordset.MoveLast
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
  '[ Finalizo la conexi�n a la base de datos ]
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
   Case 0     ' centro de costo
    tdbHelp.Columns(0).DataField = "codcco": tdbHelp.Columns(1).DataField = "detcco"
    ' Recupero la informaci�n
    s_Sql = gdl_Funcion.HelpTablas("cco", tdbHelp.Columns(0).DataField, pn_NivelCenCosto, "")
  End Select
  ' Recupera informaci�n
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fDistribucionCenCosto.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fDistribucionCenCosto.dcaRegistro.Recordset.BOF Then fDistribucionCenCosto.dcaRegistro.Recordset.MovePrevious
           If fDistribucionCenCosto.dcaRegistro.Recordset.BOF Then fDistribucionCenCosto.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fDistribucionCenCosto.dcaRegistro.Recordset.EOF Then fDistribucionCenCosto.dcaRegistro.Recordset.MoveNext
           If fDistribucionCenCosto.dcaRegistro.Recordset.EOF Then fDistribucionCenCosto.dcaRegistro.Recordset.MoveLast
   Case 3: fDistribucionCenCosto.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Periodo As String, s_Personal As String
  Dim n_Porcentaje As Double
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCenCosto.Text = "" Then Beep: MsgBox "Debe Ingresar Centro de Costo" & lblTitle, vbExclamation: txtCenCosto.SetFocus: Exit Sub
  If lblHelp(0).Caption = "???" Then Beep: MsgBox "Centro Costo " & lblTitle & " no es valido; Verificar", vbExclamation: txtCenCosto.SetFocus: Exit Sub
  If CDec(txtPorcentaje.Text) <= 0 Then Beep: MsgBox "Debe Ingresar Porcentaje de la " & lblTitle, vbExclamation: txtPorcentaje.SetFocus: Exit Sub
  ' Capturo el registro a actualizar
  s_CentroCosto = txtCenCosto.Text
  s_Periodo = o_SelDisCenCosto.txtPeriodo.Text
  s_Personal = o_SelDisCenCosto.dcaRegistro.Recordset!codpsn
  ' Validaci�n limite 100
  s_Sql = "SELECT IFNULL(ROUND(SUM(porcentaje), 2), 0) AS porcentaje "
  s_Sql = s_Sql & "FROM plcencospro "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo='" & s_Periodo & "' "
  s_Sql = s_Sql & "AND codpsn='" & s_Personal & "' "
  s_Sql = s_Sql & "AND codcco<>'" & s_CentroCosto & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  n_Porcentaje = CDec(porstRecordset!porcentaje)
  If (CDec(txtPorcentaje.Text) + n_Porcentaje) > 100 Then Beep: MsgBox "Porcentaje el Valor de " & lblTitle, vbExclamation: txtPorcentaje.SetFocus: Exit Sub
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualizaci�n
  a_Campos = Array("codcls", "codpdo", "codpsn", "codcco", "porcentaje", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, s_Periodo, s_Personal, txtCenCosto.Text, CDec(txtPorcentaje.Text), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpdo", "codpsn", "codcco")
  
  '[ Inicio la conexi�n a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacci�n
  ' Realizo el proceso de actualizaci�n de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plcencospro", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plcencospro", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacci�n
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fDistribucionCenCosto.dcaRegistro, fDistribucionCenCosto.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fDistribucionCenCosto.dcaRegistro.Recordset.Find ("codcco='" & s_CentroCosto & "'")
  ' si es actualizaci�n pasa al modo visualizaci�n
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtCenCosto.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexi�n a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminaci�n
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posici�n y titulo del formulario
  Me.Height = 4260: Me.Width = 7290
  Me.Left = 3980: Me.Top = 1830
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualizaci�n Distribucci�n por Centro de Costo"
  lblTitle = "Distribuci�n Centro Costo"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operaci�n del registro
  Me.Tag = fDistribucionCenCosto.Tag
  
  ' Configuro parametros de visualizaci�n del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y t�tulo del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "A�adir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualizaci�n del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y t�tulo del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualizaci�n
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Informaci�n de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Informaci�n de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fDistribucionCenCosto.dcaRegistro.Recordset.EOF Or fDistribucionCenCosto.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fDistribucionCenCosto.dcaRegistro.Recordset!codcco
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuraci�n de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "C�digo", "Descripci�n")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Planilla", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
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

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then Beep: MsgBox "No existen Registros para Seleccionar", vbExclamation: Exit Sub
  Select Case n_IndexHelp
   Case 0       ' centro de costo
    txtCenCosto.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtCenCosto.SetFocus
  End Select
   
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la informaci�n ordenada
  Select Case n_IndexHelp
   Case 0     ' centro de costo
    s_Sql = gdl_Funcion.HelpTablas("cco", tdbHelp.Columns(ColIndex).DataField, pn_NivelCenCosto, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtCenCosto_GotFocus()
  gdl_Procedure.MarcaGet txtCenCosto
End Sub
Private Sub txtCenCosto_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If txtCenCosto.Text = "" Then Beep: MsgBox "Debe Ingresar el C�digo del " & lblTitle, vbExclamation: txtCenCosto.SetFocus
    SendKeys "{TAB}"
    KeyAscii = 0
  End If

End Sub
Private Sub txtCenCosto_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, ps_CodEmpresa, txtCenCosto.Text, "CC")
End Sub
Private Sub txtPorcentaje_GotFocus()
  gdl_Procedure.MarcaGet txtPorcentaje
End Sub
Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    If CDec(txtPorcentaje.Text) <= 0 Then Beep: MsgBox "Debe Ingresar el Valor de la " & lblTitle, vbExclamation: txtPorcentaje.SetFocus
    SendKeys "{TAB}"
    KeyAscii = 0
  End If

End Sub
Private Sub txtPorcentaje_Validate(Cancel As Boolean)
  txtPorcentaje.Text = IIf(Not IsNumeric(txtPorcentaje.Text), 0, txtPorcentaje.Text)
  txtPorcentaje.Text = FormatNumber(CDec(txtPorcentaje.Text), 2)
End Sub
