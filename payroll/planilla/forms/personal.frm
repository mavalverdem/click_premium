VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "personal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   8865
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "personal.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "personal.frx":0166
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   10
      Top             =   600
      Width           =   7960
      _ExtentX        =   14049
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
   Begin MSAdodcLib.Adodc dcaRegistro 
      Height          =   330
      Left            =   45
      Top             =   5505
      Width           =   7960
      _ExtentX        =   14049
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
      Height          =   5235
      Index           =   0
      Left            =   8065
      TabIndex        =   0
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9234
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
         TabIndex        =   2
         Tag             =   "0"
         Top             =   1125
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
         Picture         =   "personal.frx":02C0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1545
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
         Picture         =   "personal.frx":02DC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Tag             =   "0"
         Top             =   2310
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
         Picture         =   "personal.frx":02F8
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   2745
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
         Picture         =   "personal.frx":0314
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   3900
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
         Picture         =   "personal.frx":0330
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   4335
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
         Picture         =   "personal.frx":034C
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   9
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
         TabIndex        =   1
         Tag             =   "0"
         Top             =   690
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
         Picture         =   "personal.frx":0368
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   3165
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
         Picture         =   "personal.frx":0384
      End
   End
   Begin MSComctlLib.Toolbar toolbarexp 
      Height          =   570
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1005
      ButtonWidth     =   2275
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar a Excel"
            Key             =   "Exportar a Excel"
            Object.ToolTipText     =   "Exportar Registro de Documentos a Excel"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   25
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A2"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A3"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A4"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A5"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A6"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A7"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A8"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A9"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A10"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A11"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A12"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A13"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A14"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A15"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A16"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A17"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A18"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A19"
               EndProperty
               BeginProperty ButtonMenu20 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A20"
               EndProperty
               BeginProperty ButtonMenu21 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A21"
               EndProperty
               BeginProperty ButtonMenu22 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A22"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu23 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A23"
               EndProperty
               BeginProperty ButtonMenu24 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A24"
               EndProperty
               BeginProperty ButtonMenu25 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A25"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesos"
            Key             =   "ProcesosA"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B2"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B3"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Threed.SSRibbon ribParametro 
      Height          =   360
      Index           =   1
      Left            =   7815
      TabIndex        =   12
      Top             =   120
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
      PictureUp       =   "personal.frx":03A0
   End
   Begin Threed.SSRibbon ribParametro 
      Height          =   360
      Index           =   0
      Left            =   7410
      TabIndex        =   13
      Top             =   120
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
      PictureUp       =   "personal.frx":03BC
   End
   Begin Threed.SSRibbon ribParametro 
      Height          =   360
      Index           =   2
      Left            =   8220
      TabIndex        =   14
      Top             =   120
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
      PictureUp       =   "personal.frx":03D8
   End
End
Attribute VB_Name = "fPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String                      ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, i As Integer, n_Index As Integer             ' Indice de la barra de herramientas, indice para bucle
Dim cnn As ADODB.Connection
Dim Val As String

Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, codpsn, apepaterno, apematerno, nombres, "
  s_Sql = s_Sql & "fecnacimiento, ubigeonac, nacionalidad, codpemi, naciextrapsn, sexopsn, codniv, "
  s_Sql = s_Sql & "refedirec, codvia, nomviadirec, numerdirec, "
  s_Sql = s_Sql & "intedirec, codzona, nomzondirec, ubigeodir, "
  s_Sql = s_Sql & "estcivilpsn, numhijo, numdepen, coddci, numdociden, "
  s_Sql = s_Sql & "numdocmil, codldn, telefono, celular, dctojudicial, pordsctojudi, "
  s_Sql = s_Sql & "fecingreso, reingreso, jornadalaboral, codtpt, codcgo, cgoconfianza, codpfs, "
  s_Sql = s_Sql & "codcco, codcdt, codafp, numeroafp, afpmixta, pagodolar, periodicidad, tippago, codbcopago, "
  s_Sql = s_Sql & "cuentapago, interbankpago, codbnkpago, ctsdolar, ctsdeposito, codbcocts, cuentacts, interbankcts, codbnkcts, cuentaibankcts, "
  s_Sql = s_Sql & "codeps, regpension, fecingregpen, essvida, cobsctr, afilsindical, "
  s_Sql = s_Sql & "remimprecisa, remintegralgrati, remintegralvaca, remintegralcts, "
  s_Sql = s_Sql & "remuneta, netocpc, variacpc, imporemuneto, fecbaja, nroessalud, "
  s_Sql = s_Sql & "codubica, codsec, coddeudor, codacredor, fecestado, fotopsn, estadopsn,correoelect, "
  s_Sql = s_Sql & "chkSCTRP, chkRL, chkDIS, chkMAX, chkREG, chkNOC, chkQUI,chkOIQ,siteps, segmedico,resfamiliar,ifnull(forprofesional,1) as forprofesional,finperiodo,modformativa, chkpe,cmbcatocupacional,cmbtributacion,chk27252 "
  s_Sql = s_Sql & "FROM plpersonal "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & " AND estadopsn" & IIf(ribParametro(1).Value, "<>'I'", "='I'")
  End If
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim sDireccion As String, sRegPatronal As String
  
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0, 2  ' Visualizar o analizar, eliminar registro
    If Not (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Then
      Me.Tag = IIf(Index = 0, s_MdoData_Vis, s_MdoData_Del)
      fAbcPersonal.lblTitle = "Trabajador(a) " & " " & tdbRegistro.Columns(2).Text & " " & tdbRegistro.Columns(3).Text & " " & tdbRegistro.Columns(4).Text

      fAbcPersonal.Show
      
    End If
   Case 1 ' Nuevo registro
    Me.Tag = s_MdoData_Ins
    fAbcPersonal.Show
   Case 3, 4  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index - 2, " ASC", " DESC")
   Case 5 ' Busqueda de registro
    If Not (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Then
      Set go_tdbBusqueda = tdbRegistro
      Set go_dcaBusqueda = dcaRegistro
      gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
      fBusqueda.Show vbModal
    End If
   Case 6, 7  ' Opciones de impresión
    ' Verifico que Existan Registros
    If dcaRegistro.Recordset.RecordCount = 0 Then Beep: MsgBox "No Existen " & s_TitleTable & " para Imprimir", vbExclamation: Exit Sub
    
    ' Obtengo los datos d ela empresa
    sDireccion = "": sRegPatronal = ""
    s_Sql = "SELECT via.abrevia, cfg.direccionvia, cfg.numerodir, zon.abrezona, cfg.direccionzona, cfg.ubigeodir, cfg.regpatronal "
    s_Sql = s_Sql & "FROM plcfgempresa cfg "
    s_Sql = s_Sql & "LEFT JOIN pltipovia via ON cfg.codvia=via.codvia "
    s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON cfg.codzona=zon.codzona "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sRegPatronal = gdl_Funcion.aTexto(porstRecordset!regpatronal)
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDireccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!direccionzona) & " - " & sDireccion
    End If
    porstRecordset.Close
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "PADRON GENERAL DE PERSONAL"
    gdl_Procedure.ps_ReportName = "lstpadronpsn"
    ReDim aElemento(3, 5): ReDim aElementos(2)
    ' Parametros del store procedure
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = "": aElemento(0, 2) = ""
    aElemento(0, 3) = "": aElemento(0, 4) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Direccion;" & sDireccion & ";true"
    aElemento(2, 3) = "NroRuc;" & ps_RucEmpresa & ";true"
    aElemento(2, 4) = "RegPatronal;" & sRegPatronal & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "SELECT psn.codpsn, "
    s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS apellidopsn, "
    s_Sql = s_Sql & "IFNULL(psn.nombres, '') AS nombrepsn, "
    s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, cgo.descgo, "
    s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(via.abrevia, '')), ' ', TRIM(IFNULL(psn.nomviadirec, '')), ' ', TRIM(IFNULL(psn.numerdirec, '')), ' - ', TRIM(IFNULL(psn.intedirec, '')), ' ', TRIM(IFNULL(zon.abrezona, '')), ' ', TRIM(IFNULL(psn.nomzondirec, ''))) AS direccion, "
    s_Sql = s_Sql & "TRIM(IFNULL(ubg.desubg, '')) AS distrito, "
    s_Sql = s_Sql & "psn.nacionalidad, psn.fecnacimiento, IF(psn.sexopsn='0', 'M', 'F') AS sexopsn, psn.numdociden, psn.estcivilpsn, "
    s_Sql = s_Sql & "afp.desafp, psn.numeroafp, eps.deseps, psn.nroessalud, psn.estadopsn "
    s_Sql = s_Sql & "FROM plpersonal psn "
    s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
    s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
    s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
    s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
    s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
    s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON psn.ubigeodir=ubg.codubg AND nivelubg='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "ORDER BY codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' ]
   
  End Select

End Sub
Private Sub dcaRegistro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  If FormVisible("fAbcPersonal") Then
    If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
      fAbcPersonal.ShowScreen
    End If
  End If

End Sub
Private Sub Form_Load()

  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
  cnn.CursorLocation = adUseClient
  cnn.Open

  Dim Item As New ValueItem

'Adicionar nombre al toolbar
  
  toolbarexp.Buttons(1).ButtonMenus(1).Text = "Datos Personales"
  toolbarexp.Buttons(1).ButtonMenus(2).Text = "Datos de Domicilio"
  toolbarexp.Buttons(1).ButtonMenus(3).Text = "Datos de la Empresa"
  toolbarexp.Buttons(1).ButtonMenus(4).Text = "Remuneraciones"
  toolbarexp.Buttons(1).ButtonMenus(5).Text = "Experiencia Laboral"
  toolbarexp.Buttons(1).ButtonMenus(6).Text = "Estudios Realizados"
  toolbarexp.Buttons(1).ButtonMenus(7).Text = "Contratos de Trabajo"
  toolbarexp.Buttons(1).ButtonMenus(8).Text = "Datos Familiares"
  toolbarexp.Buttons(1).ButtonMenus(9).Text = "Ingresos Descuentos Anteriores"
  toolbarexp.Buttons(1).ButtonMenus(10).Text = "Subsidios / No Laborados / No Subsidiados"
  toolbarexp.Buttons(1).ButtonMenus(11).Text = "Otros Empleadores"
  toolbarexp.Buttons(1).ButtonMenus(12).Text = "Establecimientos donde labora el Trabajador"
  toolbarexp.Buttons(1).ButtonMenus(13).Text = "Información Personal Básica "
  toolbarexp.Buttons(1).ButtonMenus(14).Text = "Planilla de Remuneraciones"
  toolbarexp.Buttons(1).ButtonMenus(16).Text = "Listado de Conceptos"
  toolbarexp.Buttons(1).ButtonMenus(17).Text = "Listado de Centro de Costos"
  toolbarexp.Buttons(1).ButtonMenus(18).Text = "Listado de AFPs"
  toolbarexp.Buttons(1).ButtonMenus(19).Text = "Listado de Bancos"
  toolbarexp.Buttons(1).ButtonMenus(20).Text = "Listado de Cargos"
  toolbarexp.Buttons(1).ButtonMenus(21).Text = "Listado de Niveles de Estudio"
  toolbarexp.Buttons(1).ButtonMenus(23).Text = "Planilla General Ingresos " & ps_Anyo
  toolbarexp.Buttons(1).ButtonMenus(24).Text = "Planilla General Descuentos " & ps_Anyo
  toolbarexp.Buttons(1).ButtonMenus(25).Text = "Planilla General Aportaciones " & ps_Anyo
  toolbarexp.Buttons(3).ButtonMenus(1).Text = "Establecimientos por Trabajador"
  toolbarexp.Buttons(3).ButtonMenus(2).Text = "Actualizar datos Subsidios/No Laborados/No Subsidiados"
  toolbarexp.Buttons(3).ButtonMenus(3).Text = "Actualizar Prorrateo de Centro de Costos"
  ' Establece posición del formulario
  Me.Height = 6315: Me.Width = 8950
  Me.Left = 105: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Trabajador(es)"
  
  ReDim aElemento(7, 10)
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 0) = Choose(i + 1, "Código", "Numero Doc.", "Apellido Paterno", "Apellido Materno", "Nombre(s)", "C.Costo", "Ok")
    aElemento(i, 1) = Choose(i + 1, "codpsn", "numdociden", "apepaterno", "apematerno", "nombres", "codcco", "estadopsn")
    aElemento(i, 2) = Choose(i + 1, 980, 900, 1550, 1550, 1320, 800, 300)
    aElemento(i, 3) = Choose(i + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(i, 4) = Choose(i + 1, "", "", "", "", "", "", "")
    aElemento(i, 5) = Choose(i + 1, False, False, False, False, False, False, False)
    aElemento(i, 6) = Choose(i + 1, True, True, True, True, True, True, True)
    aElemento(i, 7) = Choose(i + 1, "", "", "", "", "", "", "")
    aElemento(i, 8) = Choose(i + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(i, 9) = Choose(i + 1, 0, 0, 0, 0, 0, 0, 0)
  Next i
  ReDim aElementos(1, 3)
  For i = 0 To (UBound(aElementos, 1) - 1)
    aElementos(i, 0) = ""
    aElementos(i, 1) = 13427690: aElementos(i, 2) = vbBlack
  Next i
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(6).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(6).ValueItems.Translate = True
  For i = 0 To 5
    tdbRegistro.Columns(6).ValueItems.Add Item
    tdbRegistro.Columns(6).ValueItems.Item(i).Value = Choose(i + 1, "A", "V", "L", "P", "O", "I")
    tdbRegistro.Columns(6).ValueItems.Item(i).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(i% + 1, "estadok", "estadovo", "estadnok", "estadopk", "estadopn", "procenok") & ".bmp")
  Next i
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(8, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "registro": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  
  
  ' Cargo los graficos a los controles
  For i = 0 To (UBound(aElemento, 1) - 1)
      aElemento(i, 1) = Choose(i + 1, "seleccio", "anadir", "borrar", "ordascen", "orddesce", "busqueda", "prelimin", "Imprimir")
      aElemento(i, 2) = Choose(i + 1, "Selecciona y Edita " & s_TitleTable, "Añadir " & s_TitleTable, "Eliminar " & s_TitleTable, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Presentación Preliminar", "Imprimir")
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  
  ribParametro(0).Value = True
  
End Sub
Private Sub Form_Unload(Cancel As Integer)

  If FormVisible("fAbcPersonal") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcPersonal.Caption, vbExclamation
    Cancel = True
  End If

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
 RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
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
Private Sub toolbarexp_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim mesp As String, anop As String
Dim rsselect As New Recordset
Dim rsdelete As New Recordset
Dim rsinsert As New Recordset
Dim s_sql_select As String
Dim s_sql_delete As String
Dim s_sql_insert As String

Select Case ButtonMenu.Key
Case "A1"
  exportar ("DP")
Case "A2"
    exportar ("DD")
Case "A3"
    exportar ("DE")
Case "A4"
    exportar ("RE")
Case "A5"
    exportar ("EL")
Case "A6"
    exportar ("ER")
Case "A7"
    exportar ("CT")
Case "A8"
    exportar ("DF")
Case "A9"
    exportar ("ID")
Case "A10"
    exportar ("FA")
Case "A11"
    exportar ("OE")
Case "A12"
    exportar ("ET")
Case "A13"
    ppExcelInfoBasica
Case "A14"
    Val = InputBox("Ingrese Mes Ej. 01", "Planilla de Remuneraciones", "01")
    If (CInt(Val) <= 1 And CInt(Val) <= 12) Then
    Else
      Exit Sub
    End If
    exportar ("PL")
Case "A16"
    exportar ("CO")
Case "A17"
    exportar ("CC")
Case "A18"
    exportar ("AP")
Case "A19"
    exportar ("BA")
Case "A20"
    exportar ("CA")
Case "A21"
    exportar ("NE")
Case "A23"
    exportar ("PGI")
Case "A24"
    exportar ("PGD")
Case "A25"
    exportar ("PGA")
Case "B1"
  Val = InputBox("Ingrese Mes Ej. 01 " & vbCrLf & vbCrLf & "Se Eliminara los Establecimientos del Mes y Año a Procesar", "Planilla de Remuneraciones", "01")
  If (CInt(Val) <= 1 And CInt(Val) <= 12) Then
    If Val = "01" Then
      mesp = "12"
      anop = CInt(ps_Anyo) - 1
    Else
      mesp = CInt(Val) - 1
      anop = ps_Anyo
      If Len(mesp) = 1 Then mesp = "0" & mesp
    End If
    TipodeProgreso = 1
    IntervalodeTiempo = 100
    labelprogreso = "Generando                                                          Datos"
    s_sql_select = "SELECT codpsn FROM plestalaboral where ano='" & anop & "' and mes='" & mesp & "' and codcls='" & ps_ClsPlanilla & "' group by codcls,codpsn "
    rsselect.Open s_sql_select, cnn, adOpenStatic, adLockOptimistic
    On Error GoTo Errores1
    rsselect.MoveFirst
    For i = 1 To rsselect.RecordCount
        s_sql_delete = "delete from plestalaboral where ano='" & ps_Anyo & "' and mes='" & Val & "' and codcls='" & ps_ClsPlanilla & "' and codpsn='" & rsselect(0) & "'"
        rsdelete.Open s_sql_delete, cnn, adOpenStatic, adLockOptimistic
        On Error GoTo Errores1
        s_sql_insert = " insert into plestalaboral (codcls,codpsn,orden,ano,mes,ruc,codest,tasa,usrcre,fyhcre)"
        s_sql_insert = s_sql_insert & " select est.codcls,est.codpsn,est.orden,'" & ps_Anyo & "','" & Val & "',est.ruc,est.codest,est.tasa,'admin',now() from plestalaboral est left join plpersonal psn on est.codcls=psn.codcls and est.codpsn=psn.codpsn   "
        s_sql_insert = s_sql_insert & " where psn.estadopsn <> 'I' and est.ano='" & anop & "' and est.mes='" & mesp & "' and est.codcls='" & ps_ClsPlanilla & "' and est.codpsn='" & rsselect(0) & "'"
        rsinsert.Open s_sql_insert, cnn, adOpenStatic, adLockOptimistic
        On Error GoTo Errores1
        rsselect.MoveNext
    Next
    Progreso.Show vbModal
Errores1:     Exit Sub
  Else
    Exit Sub
  End If
Case "B2"

Val = InputBox("Ingrese Mes Ej. 01 " & vbCrLf & vbCrLf & "Se copiara los datos de Vacaciones de Asistencias a Tabla Subsidios/No Laborados/No Subsidiados del Mes y Año a Procesar", "Planilla de Remuneraciones", "01")
If (CInt(Val) <= 1 And CInt(Val) <= 12) Then
    mesp = Val
    anop = ps_Anyo
    
    TipodeProgreso = 1
    IntervalodeTiempo = 100
    labelprogreso = "Generando                                                          Datos"
    s_sql_select = "select a.codpsn from plasistencia a inner join plperiodo p on a.codpdo=p.codpdo and a.codcls=p.codcls where a.codcls='" & ps_ClsPlanilla & "' and p.mespdo='" & mesp & "' and p.anopdo='" & anop & "' and a.diavacaciones>0 "
    rsselect.Open s_sql_select, cnn, adOpenStatic, adLockOptimistic
    On Error GoTo Errores2
    rsselect.MoveFirst
    For i = 1 To rsselect.RecordCount
        s_sql_delete = "delete from plsubsidios where ano='" & anop & "' and mes='" & mesp & "' and codcls='" & ps_ClsPlanilla & "' and codpsn='" & rsselect(0) & "' and tipsub='23' "
        rsdelete.Open s_sql_delete, cnn, adOpenStatic, adLockOptimistic
        On Error GoTo Errores2
        rsselect.MoveNext
    Next
    s_sql_insert = " insert into plsubsidios (codcls,codpsn,orden,ano,mes,citsub,tipsub,fechaini,fechafin,usrcre,fyhcre) "
    s_sql_insert = s_sql_insert & " select a.codcls,a.codpsn,"
    s_sql_insert = s_sql_insert & " coalesce((select max(s.orden)+1 from plsubsidios s where s.codcls='" & ps_ClsPlanilla & "' and s.codpsn=a.codpsn and s.mes='" & mesp & "' and s.ano='" & anop & "'),1) as orden,"
    s_sql_insert = s_sql_insert & " p.anopdo,p.mespdo,'VAC' as citsub,'23' as tipsub,a.fechainivaca1,a.fechafinvaca1,'admin',now()  from plasistencia a"
    s_sql_insert = s_sql_insert & " inner join plperiodo p on a.codpdo=p.codpdo and a.codcls=p.codcls"
    s_sql_insert = s_sql_insert & " where a.codcls='" & ps_ClsPlanilla & "' and p.mespdo='" & mesp & "' and p.anopdo='" & anop & "' and a.diavacaciones>0 "
    rsinsert.Open s_sql_insert, cnn, adOpenStatic, adLockOptimistic
    On Error GoTo Errores2
    Progreso.Show vbModal
Errores2:     Exit Sub
Else
    Exit Sub
End If
Case "B3"

    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
        tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
        MsgBox tdbRegistro.Columns(0).Text
    Next n_Index
    
    Exit Sub
    
    Dim rpta As Integer
    rpta = MsgBox("Va a Copiar Centro de Costos del Registro de Personal a Centro de Costos Prorrateo", vbQuestion + vbYesNo + vbDefaultButton2, "Salir")
    If rpta = 6 Then
    TipodeProgreso = 1
    IntervalodeTiempo = 100
    labelprogreso = "Generando                                                          Datos"
    s_sql_insert = " insert into plcencospro (codcls,codpsn,codcco,porcentaje,usrcre,fyhcre) "
    s_sql_insert = s_sql_insert & " select codcls,codpsn,codcco,100,'admin',now() from plpersonal "
    s_sql_insert = s_sql_insert & " where codcls='" & ps_ClsPlanilla & "' and estadopsn='A' and codpsn not in (select codpsn from plcencospro where codcls='" & ps_ClsPlanilla & "' group by codpsn)"
    rsinsert.Open s_sql_insert, cnn, adOpenStatic, adLockOptimistic
    On Error GoTo Errores3
    Progreso.Show vbModal
Errores3:     Exit Sub
    Else
    Exit Sub
    End If
End Select
End Sub
Private Sub ppExcelInfoBasica()
  Dim sHojaExcel As String, sPersonal As String, sConcepto As String
  Dim nSecuencia As Long, nColumna As Long, nTitulo As Long
  Dim poApExcel As Object
  Dim a_Ingreso()
    
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroEnEspera
  
  sHojaExcel = "Informa_basica"
  
  s_Sql = "SELECT psn.codcls, psn.codpsn, psn.numdociden, cco.detcco, cls.descls, "
  s_Sql = s_Sql & "CONCAT(psn.apepaterno, ' ', psn.apematerno) AS apellidos, psn.nombres, DATE_FORMAT(psn.fecingreso,'%d/%m/%Y') AS fecingreso, "
  s_Sql = s_Sql & "tco.destco, DATE_FORMAT((CASE WHEN con.tipcon='01' THEN Null ELSE con.fechafin END),'%d/%m/%Y') AS fechafin, cdt.descdt, cgo.descgo, "
  s_Sql = s_Sql & "afp.desafp, psn.afpmixta as tipo_comision,DATE_FORMAT(psn.fecnacimiento,'%d/%m/%Y') AS fecnacimiento, "
  s_Sql = s_Sql & "(CASE WHEN rmd.codmon='" & s_Codmon_mn & "' THEN '" & s_Codmon_mn_Txt & "' ELSE '" & s_Codmon_me_Txt & "' END) AS codmon, "
  s_Sql = s_Sql & "cxp.defaultcpc, rmd.codcpc, cpc.aliascpc, rmd.imporemune "
  s_Sql = s_Sql & "FROM plpersonal psn "
  s_Sql = s_Sql & "INNER JOIN plclasplan cls ON cls.codcls=psn.codcls "
  s_Sql = s_Sql & "LEFT JOIN plremudefa rmd ON rmd.codcls=psn.codcls AND rmd.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON rmd.codcpc=cpc.codcpc AND cpc.tipocpc='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "INNER JOIN plconceplanilla cxp ON cxp.codcls=rmd.codcls AND cxp.codcpc=rmd.codcpc "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON afp.codafp=psn.codafp "
  s_Sql = s_Sql & "INNER JOIN plcargo cgo ON cgo.codcls=psn.codcls AND cgo.codcgo=psn.codcgo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON cco.codcco=psn.codcco "
  s_Sql = s_Sql & "LEFT JOIN plconditrabajo cdt ON cdt.codcls=psn.codcls AND cdt.codcdt=psn.codcdt "
  s_Sql = s_Sql & "LEFT JOIN plcontrato con ON con.codcls=psn.codcls AND con.codpsn=psn.codpsn AND con.estadocon='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "LEFT JOIN pltipcontrato tco ON tco.codtco=con.tipcon "
  s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND psn.estadopsn<>'I' "
  s_Sql = s_Sql & "ORDER BY psn.codpsn, rmd.codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  On Error GoTo Finalizar
  'IntervalodeTiempo = porstRecordset.RecordCount
  porstRecordset.MoveFirst
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    Set poApExcel = CreateObject("Excel.application")
    poApExcel.Visible = False
    poApExcel.Workbooks.Add
    poApExcel.Sheets("Hoja1").Name = sHojaExcel
    
    poApExcel.ActiveWindow.Zoom = 75
    poApExcel.Cells(1, 1).Formula = "Información Basica de Personal"
    poApExcel.Cells(1, 1).Font.Size = 18
    poApExcel.Cells(2, 1).Formula = ""

    ' Inicializo variables
    ReDim a_Ingreso(3, 1)
    a_Ingreso(1, 1) = porstRecordset!codcpc
    a_Ingreso(2, 1) = UCase(porstRecordset!aliascpc)
    a_Ingreso(3, 1) = 0
    
    nTitulo = 15
    For nColumna = 1 To nTitulo
      poApExcel.Cells(4, nColumna).Formula = Choose(nColumna, "Nº", "CODIGO", "DOCUMENTO IDENTIDAD", "CENTRO COSTO", "PLANILLA", "APELLIDOS", "NOMBRES", "FECHA INGRESO", "TIPO CONTRATO", "VCTO CONTRATO", "MODALIDAD CONTRATO", "CARGO", "SISTEMA PENSIÓN", "AFP-COMISION MIXTA", "FECHA NACIMIENTO", "MONEDA")
    Next nColumna
    While Not porstRecordset.EOF
      ' Inicializo conceptos
      For nColumna = 1 To UBound(a_Ingreso, 2): a_Ingreso(3, nColumna) = 0: Next nColumna
      sPersonal = porstRecordset!codpsn
      Do
        For nColumna = 1 To UBound(a_Ingreso, 2)
          If a_Ingreso(1, nColumna) = porstRecordset!codcpc Then Exit For
        Next nColumna
        If nColumna > UBound(a_Ingreso, 2) Then
          ReDim Preserve a_Ingreso(3, nColumna)
          a_Ingreso(1, nColumna) = porstRecordset!codcpc
          a_Ingreso(2, nColumna) = UCase(porstRecordset!aliascpc)
        End If
        a_Ingreso(3, nColumna) = CDec(porstRecordset!imporemune)
        ' siguiente registro
        porstRecordset.MoveNext
        If porstRecordset.EOF Then Exit Do
      Loop While (sPersonal = porstRecordset!codpsn)
      porstRecordset.MovePrevious
      ' Información de personal
      nSecuencia = nSecuencia + 1
      For nColumna = 1 To nTitulo
        poApExcel.Cells(nSecuencia + 4, nColumna).Formula = IIf(nColumna = 1, nSecuencia, porstRecordset(nColumna - 1))
      Next nColumna
      For nColumna = 1 To UBound(a_Ingreso, 2)
        poApExcel.Cells(nSecuencia + 4, (nTitulo + nColumna)).Formula = a_Ingreso(3, nColumna)
      Next nColumna
      ' siguiente registro
      porstRecordset.MoveNext
    Wend
    ' Titulo de conceptos
    For nColumna = 1 To UBound(a_Ingreso, 2)
      poApExcel.Cells(4, (nTitulo + nColumna)).Formula = a_Ingreso(2, nColumna)
    Next nColumna
    ' Remuneración total
    poApExcel.Cells(4, (nTitulo + nColumna)).Formula = "TOTAL REMUNERACIÓN"
    nColumna = nTitulo + nColumna
    For nTitulo = 1 To nSecuencia
      poApExcel.Cells(nTitulo + 4, nColumna).FormulaR1C1 = "=ROUND(SUM(RC[-" & UBound(a_Ingreso, 2) & "]:RC[-1]),2)"
    Next nTitulo
  End If
  TipodeProgreso = 1
  IntervalodeTiempo = 100
  labelprogreso = "Exportando Datos a Excel"
  Progreso.Show vbModal
  
  MsgBox ("Proceso de Exportacion a Excel, terminado")
  poApExcel.Visible = True

Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
End Sub
Private Sub exportar(table As String)
Dim nhoja As String
Dim strsql As String
Dim i As Integer
Dim j As Integer
Dim cols As Integer
Select Case table
Case "DP"
    nhoja = "Datos Personales"
    strsql = "select codcls,codpsn,apepaterno,apematerno,nombres,fecnacimiento,ubigeonac,nacionalidad,naciextrapsn,sexopsn,refedirec,codvia,nomviadirec,numerdirec,intedirec,codzona,nomzondirec,ubigeodir,estcivilpsn,numhijo,numdepen,coddci,numdociden,numdocmil,telefono,celular,dctojudicial,pordsctojudi,fecingreso,codtpt,codcgo,cgoconfianza,codpfs,codcco,codafp,numeroafp,afpmixta,pagodolar,codbcopago,cuentapago,ctsdeposito,ctsdolar,codbcocts,cuentacts,codeps,regpension,fecingregpen,essvida,cobsctr,afilsindical,remintegralgrati,remintegralvaca,remintegralcts,remimprecisa,remuneta,netocpc,variacpc,imporemuneto,fecbaja,nroessalud,codubica,codsec,coddeudor,codacredor,fecestado,fotopsn,estadopsn,usrcre,fyhcre,usrmdf,fyhmdf,correoelect,chkSCTRP,"
    strsql = strsql & "chkRL,chkDIS,chkMAX,chkREG,chkNOC,chkQUI,chkOIQ,siteps,tippago,segmedico,resfamiliar,forprofesional,finperiodo,modformativa,"
    strsql = strsql & "periodicidad , chkPE, cmbcatocupacional, cmbtributacion,chk27252 from plpersonal where codcls='" & ps_ClsPlanilla & "'"
    cols = 91
Case "DD"
    nhoja = "Datos Domiciliarios"
    strsql = "select codcls,codpsn,apepaterno,apematerno,nombres,coddci,numdociden,refedirec,codvia,nomviadirec,numerdirec,intedirec,codzona,nomzondirec,ubigeodir from plpersonal where codcls='" & ps_ClsPlanilla & "'"
    cols = 15
Case "DE"
    nhoja = "Datos Empresa"
    strsql = "select * from plpersonal"
    strsql = "select codcls,codpsn,apepaterno,apematerno,nombres,coddci,numdociden,fecingreso,codtpt,codcgo,cgoconfianza,codpfs,codcco,codafp,numeroafp,pagodolar,codbcopago,cuentapago,ctsdeposito,ctsdolar,codbcocts,cuentacts,codeps,regpension,fecingregpen,essvida,cobsctr,afilsindical,remintegralgrati,remintegralvaca,remintegralcts,remimprecisa,remuneta,netocpc,variacpc,imporemuneto,fecbaja,nroessalud,codubica,codsec,coddeudor,codacredor,fecestado,estadopsn,chkSCTRP,chkRL,chkDIS,chkMAX,chkREG,chkNOC,chkQUI,chkOIQ,siteps,tippago,segmedico,resfamiliar,forprofesional,finperiodo,modformativa,periodicidad from plpersonal where codcls='" & ps_ClsPlanilla & "'"
    cols = 60
Case "RE"
    nhoja = "Remuneraciones"
    strsql = "select coddci,numdociden,plremudefa.codcls,plremudefa.codpsn,codcpc,codmon,imporemune from plremudefa inner join plpersonal on plremudefa.codpsn= plpersonal.codpsn and plremudefa.codcls= plpersonal.codcls where plremudefa.codcls='" & ps_ClsPlanilla & "'"
    cols = 7
Case "EL"
    nhoja = "Experiencia Laboral"
    strsql = "select coddci,numdociden,plexpelaboral.codcls,plexpelaboral.codpsn,orden,empresa,plexpelaboral.codcgo,fechaini,fechafin,observacion from plexpelaboral inner join plpersonal on plexpelaboral.codpsn= plpersonal.codpsn and plexpelaboral.codcls= plpersonal.codcls where plexpelaboral.codcls='" & ps_ClsPlanilla & "'"
    cols = 10
Case "ER"
    nhoja = "Estudios Realizados"
    strsql = "select coddci,numdociden,plestudios.codcls,plestudios.codpsn,orden,institucion,grado,fechaini,fechafin,observacion from plexpelaboral inner join plpersonal on plestudios.codpsn= plpersonal.codpsn and plestudios.codcls= plpersonal.codcls where plestudios.codcls='" & ps_ClsPlanilla & "'"
    strsql = "select * from plestudios"
    cols = 12
Case "CT"
    nhoja = "Contrato de Trabajo"
    strsql = "select coddci,numdociden,plcontrato.codcls,plcontrato.codpsn,numdocumen,ano,mes,dia,fechaini,fechafin,observacion,archivo,estadocon,tipcon from plcontrato inner join plpersonal on plcontrato.codpsn= plpersonal.codpsn and plcontrato.codcls= plpersonal.codcls where plcontrato.codcls='" & ps_ClsPlanilla & "'"
    cols = 14
Case "DF"
    nhoja = "Datos Familiares"
    strsql = "select plpersonal.coddci,plpersonal.numdociden,plfamiliares.codcls,plfamiliares.codpsn,orden,plfamiliares.apepaterno,plfamiliares.apematerno,plfamiliares.nombres,plfamiliares.fecnacimiento,sexofam,plfamiliares.coddci,plfamiliares.numdociden,vinculo,cartamed,domicilio,plfamiliares.codvia,plfamiliares.nomviadom,plfamiliares.numerdom,plfamiliares.intedom,plfamiliares.codzona,plfamiliares.nomzonadom,plfamiliares.refedom,plfamiliares.ubigeodom,incapacidad,certificadomed,motivoina,estadofam,tipdocpaternidad,acrepaternidad,fecalta,plfamiliares.fecbaja from plfamiliares inner join plpersonal on plfamiliares.codpsn= plpersonal.codpsn and plfamiliares.codcls= plpersonal.codcls where plfamiliares.codcls='" & ps_ClsPlanilla & "'"
    cols = 31
Case "ID"
    nhoja = "Ingresos Descuentos"
    strsql = "select coddci,numdociden,plremuexce.codcls,codpdo,plremuexce.codpsn,codcpc,codmon,imporemune from plremuexce inner join plpersonal on plremuexce.codpsn= plpersonal.codpsn and plremuexce.codcls= plpersonal.codcls where plremuexce.codcls='" & ps_ClsPlanilla & "'"
    cols = 8
Case "FA"
    nhoja = "Faltas"
    strsql = "select coddci,numdociden,plsubsidios.codcls,plsubsidios.codpsn,orden,ano,mes,citsub,fechaini,fechafin from plsubsidios inner join plpersonal on plsubsidios.codpsn= plpersonal.codpsn and plsubsidios.codcls= plpersonal.codcls where plsubsidios.codcls='" & ps_ClsPlanilla & "'"
    cols = 10
Case "OE"
    nhoja = "Otros Empleadores"
    strsql = "select coddci,numdociden,plempleadores.codcls,plempleadores.codpsn,orden,ruc,razons from plempleadores inner join plpersonal on plempleadores.codpsn= plpersonal.codpsn and plempleadores.codcls= plpersonal.codcls where plempleadores.codcls='" & ps_ClsPlanilla & "'"
    cols = 7
Case "ET"
    nhoja = "Establecimientos"
    strsql = "select coddci,numdociden,plestalaboral.codcls,plestalaboral.codpsn,orden,ano,mes,ruc,codest,tasa from plestalaboral inner join plpersonal on plestalaboral.codpsn= plpersonal.codpsn and plestalaboral.codcls= plpersonal.codcls where plestalaboral.codcls='" & ps_ClsPlanilla & "'"
    cols = 10
Case "PL"
    nhoja = "Planilla de Remuneraciones"
    strsql = "select coddci,numdociden,plresultado.codcls,codpdo,codproce,plresultado.codpsn,codcpc,secuencia,codmon,importe_mn,importe_me,codcta_debmn,codcta_habmn,codcta_debme,codcta_habme,pdoano,pdomes,tipocpc,impbolecpc,codproce_pdo from plresultado inner join plpersonal on plresultado.codpsn= plpersonal.codpsn and plresultado.codcls= plpersonal.codcls where pdoano='" & ps_Anyo & "' and pdomes='" & Val & "'  and plresultado.codcls='" & ps_ClsPlanilla & "' and impbolecpc=1"
    cols = 20
Case "CO"
    nhoja = "Listado de Conceptos"
    strsql = "select plconcepto.codcpc,descpc,aliascpc,tipocpc,obs,estadocpc from plconcepto inner join plconceplanilla on plconcepto.codcpc= plconceplanilla.codcpc where impbolecpc=1"
    cols = 6
Case "CC"
    nhoja = "Listado de Costos"
    strsql = "select codcco,detcco,detccox from cocco "
    cols = 3
Case "AP"
    nhoja = "Listado de AFPs"
    strsql = "select codafp,desafp,factor1,factor2,factor3,factor4,codbco,ctacteafp,desctacteafp,ctactefondo,desctactefondo,estadoafp from plentidadafp "
    cols = 12
Case "BA"
    nhoja = "Listado de Bancos"
    strsql = "select codbco,desbco,cuentamn,cuentame,codentidad,formato,estadobco from plbanco "
    cols = 7
Case "CA"
    nhoja = "Listado de Cargos"
    strsql = "select codcgo,descgo,estadocgo from plcargo where codcls='" & ps_ClsPlanilla & "'"
    cols = 3
Case "NE"
    nhoja = "Listado de NivEstudios"
    strsql = "select codniv,desniv,estadoniv from plniveducativo "
    cols = 3
Case "PGI"
    nhoja = "Planilla General Ingresos"
    strsql = " select codpdo,plresultado.codpsn,concat(apepaterno,' ',apematerno,' ',nombres),plresultado.codcpc,descpc,importe_mn,importe_me,pdomes,pdoano from plresultado"
    strsql = strsql & " inner join plpersonal on plresultado.codpsn=plpersonal.codpsn "
    strsql = strsql & " inner join plconcepto on plresultado.codcpc=plconcepto.codcpc "
    strsql = strsql & " Where Left(codpdo, 4) ='" & ps_Anyo & "' And impbolecpc = 1 And plresultado.tipocpc = 0 and plresultado.codcls='" & ps_ClsPlanilla & "'"
    strsql = strsql & " order by codpsn,codpdo,codcpc "
    cols = 9
Case "PGD"
    nhoja = "Planilla General Descuentos"
    strsql = " select codpdo,plresultado.codpsn,concat(apepaterno,' ',apematerno,' ',nombres),plresultado.codcpc,descpc,importe_mn,importe_me,pdomes,pdoano from plresultado"
    strsql = strsql & " inner join plpersonal on plresultado.codpsn=plpersonal.codpsn "
    strsql = strsql & " inner join plconcepto on plresultado.codcpc=plconcepto.codcpc "
    strsql = strsql & " Where Left(codpdo, 4) ='" & ps_Anyo & "' And impbolecpc = 1 And plresultado.tipocpc = 1 and plresultado.codcls='" & ps_ClsPlanilla & "'"
    strsql = strsql & " order by codpsn,codpdo,codcpc "
    cols = 9
Case "PGA"
    nhoja = "Planilla General Aportaciones"
    strsql = " select codpdo,plresultado.codpsn,concat(apepaterno,' ',apematerno,' ',nombres),plresultado.codcpc,descpc,importe_mn,importe_me,pdomes,pdoano from plresultado"
    strsql = strsql & " inner join plpersonal on plresultado.codpsn=plpersonal.codpsn "
    strsql = strsql & " inner join plconcepto on plresultado.codcpc=plconcepto.codcpc "
    strsql = strsql & " Where Left(codpdo, 4) ='" & ps_Anyo & "' And impbolecpc = 1 And plresultado.tipocpc = 2 and plresultado.codcls='" & ps_ClsPlanilla & "'"
    strsql = strsql & " order by codpsn,codpdo,codcpc "
    cols = 9
End Select
Dim rsexportar As New Recordset
Dim ApExcel As Variant
Set ApExcel = CreateObject("Excel.application")
ApExcel.Visible = False
ApExcel.Workbooks.Add
ApExcel.Sheets("Hoja1").Name = nhoja

ApExcel.ActiveWindow.Zoom = 75
ApExcel.Cells(1, 1).Formula = "Informacion del Trabajador : " & nhoja
ApExcel.Cells(1, 1).Font.Size = 18
ApExcel.Cells(2, 1).Formula = ""
'************************************
rsexportar.Open strsql, cnn, adOpenStatic, adLockOptimistic
On Error GoTo Error
IntervalodeTiempo = rsexportar.RecordCount
rsexportar.MoveFirst
For i = 1 To rsexportar.RecordCount
If i = 1 Then
    For j = 1 To cols
    ApExcel.Cells(i + 3, j).Formula = rsexportar.Fields(j - 1).Name
    Next
End If
    For j = 1 To cols
    ApExcel.Cells(i + 4, j).Formula = rsexportar(j - 1)
    Next
rsexportar.MoveNext
Next
TipodeProgreso = 1
labelprogreso = "Exportando Datos a Excel"
Progreso.Show vbModal
MsgBox ("Proceso de Exportacion a Excel, terminado")
ApExcel.Visible = True
Error:
End Sub



