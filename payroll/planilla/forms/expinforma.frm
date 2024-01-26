VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fExpInformacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "expinforma.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6135
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4605
      Left            =   45
      TabIndex        =   13
      Top             =   570
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
      Left            =   45
      Top             =   5205
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
      Height          =   4980
      Index           =   0
      Left            =   5355
      TabIndex        =   3
      Top             =   570
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8784
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
         TabIndex        =   12
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
         TabIndex        =   5
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
         Picture         =   "expinforma.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   6
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
         Picture         =   "expinforma.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   2250
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
         Picture         =   "expinforma.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2685
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
         Picture         =   "expinforma.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3795
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
         Picture         =   "expinforma.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   4230
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
         Picture         =   "expinforma.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   4
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
         Picture         =   "expinforma.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   3105
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
         Picture         =   "expinforma.frx":00D0
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
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
         ItemData        =   "expinforma.frx":00EC
         Left            =   1260
         List            =   "expinforma.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   2625
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   4500
         TabIndex        =   14
         Tag             =   "1"
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
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   4905
         TabIndex        =   15
         Tag             =   "1"
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
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   5310
         TabIndex        =   16
         Tag             =   "1"
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
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mes :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   135
         Width           =   930
      End
   End
End
Attribute VB_Name = "fExpInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(2)                               ' Array de inicio y fin de seleccion de registro
'[
Private Sub ExportaDerechoHabiente(ByVal s_Archivo As String, ByVal s_Periodo As String, s_Proceso As String, s_FechaHora As String, s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim s_Motivo As String, s_Estado As String, s_CartaMedica As String
  
  ' Recupero la información para exportar
  s_Sql = "SELECT fam.codcls, fam.codpsn, fam.orden, psn.coddci, psn.numdociden, psn.apepaterno, psn.apematerno, psn.nombres, "
  s_Sql = s_Sql & "fam.coddci AS coddcifam, dci.sigladci, fam.numdociden AS numdocidenfam, fam.apepaterno AS apepaternofam, fam.apematerno AS apematernofam, fam.nombres AS nombresfam, "
  s_Sql = s_Sql & "fam.fecnacimiento AS fecnacimientofam, fam.sexofam, fam.vinculo, fam.cartamed, "
  s_Sql = s_Sql & "fam.estadofam, fam.motivoina, fam.incapacidad, fam.certificadomed, fam.domicilio, "
  s_Sql = s_Sql & "fam.nomviadom, fam.numerdom, fam.intedom, fam.nomzonadom, fam.refedom, "
  s_Sql = s_Sql & "fam.codvia, via.desvia, fam.codzona, zon.deszona, fam.ubigeodom "
  s_Sql = s_Sql & "FROM plfamiliares fam "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON fam.codcls=psn.codcls AND fam.codpsn=psn.codpsn "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON fam.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON fam.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON fam.codzona=zon.codzona "
  s_Sql = s_Sql & "WHERE fam.codcls IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND (DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & s_Periodo & "' AND psn.estadopsn<>'I') "
  s_Sql = s_Sql & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & s_Periodo & "' AND psn.estadopsn='I') "
  s_Sql = s_Sql & "ORDER BY codcls, codpsn, orden"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando información ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    
    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("codcls", "codpsn", "orden", "coddci", "numdociden", "apepaterno", "apematerno", "nombres", "coddcifam", "numdocidenfam", "apepaternofam", "apematernofam", "nombresfam", "fecnacimientofam", "sexofam", "vinculo", "caratamed", "estadofam", "motivoina", "certificadomed", "domicilio", "nomviadom", "numerdom", "intedom", "nomzonadom", "refedom", "codvia", "codzona", "ubigeodom")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
      s_Caracter = "|"
    End If
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      s_Estado = gdl_Funcion.aTexto(porstRecordset!estadofam)
      s_Estado = IIf(s_Estado = s_Estado_Act, "10", "11")
      s_Motivo = IIf(porstRecordset!motivoina = s_Estado_Act, "2", IIf(porstRecordset!motivoina = s_Estado_Blq, "3", ""))
      s_CartaMedica = gdl_Funcion.aTexto(porstRecordset!cartamed)
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(porstRecordset!codcls, porstRecordset!codpsn, porstRecordset!orden, gdl_Funcion.aTexto(porstRecordset!coddci), gdl_Funcion.aTexto(porstRecordset!numdociden), gdl_Funcion.aTexto(porstRecordset!apepaterno), gdl_Funcion.aTexto(porstRecordset!apematerno), gdl_Funcion.aTexto(porstRecordset!nombres), gdl_Funcion.aTexto(porstRecordset!coddcifam), gdl_Funcion.aTexto(porstRecordset!numdocidenfam), gdl_Funcion.aTexto(porstRecordset!apepaternofam), gdl_Funcion.aTexto(porstRecordset!apematernofam), gdl_Funcion.aTexto(porstRecordset!nombresfam), Format(porstRecordset!fecnacimientofam, s_FmtFechMysql_0), _
                    gdl_Funcion.aTexto(porstRecordset!sexofam), gdl_Funcion.aTexto(porstRecordset!vinculo), s_CartaMedica, s_Estado, s_Motivo, gdl_Funcion.aTexto(porstRecordset!certificadomed), gdl_Funcion.aTexto(porstRecordset!Domicilio), gdl_Funcion.aTexto(porstRecordset!nomviadom), gdl_Funcion.aTexto(porstRecordset!numerdom), gdl_Funcion.aTexto(porstRecordset!intedom), gdl_Funcion.aTexto(porstRecordset!nomzonadom), gdl_Funcion.aTexto(porstRecordset!refedom), gdl_Funcion.aTexto(porstRecordset!codvia), gdl_Funcion.aTexto(porstRecordset!codzona), gdl_Funcion.aTexto(porstRecordset!ubigeodom))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = ""
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!coddci) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!numdociden) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!coddcifam) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!numdocidenfam) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!apepaternofam) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!apematernofam) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!nombresfam) & s_Caracter
        psRegistro = psRegistro & Format(porstRecordset!fecnacimientofam, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & Val(gdl_Funcion.aTexto(porstRecordset!sexofam)) + 1 & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!vinculo) & s_Caracter
        psRegistro = psRegistro & IIf(s_CartaMedica <> "", s_Estado_Act, s_CartaMedica) & s_Caracter
        psRegistro = psRegistro & s_CartaMedica & s_Caracter
        psRegistro = psRegistro & s_Estado & s_Caracter
        psRegistro = psRegistro & s_Motivo & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!certificadomed) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!Domicilio) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!nomviadom) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!numerdom) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!intedom) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!nomzonadom) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!refedom) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codvia) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codzona) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!ubigeodom) & s_Caracter
        potxtFileExp.WriteLine psRegistro
      End If
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
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
Private Sub ExportaRemuneraciones(ByVal s_Archivo As String, ByVal s_Periodo As String, s_Proceso As String, s_FechaHora As String, s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, a_Parametro(10) As String
  Dim s_Caracter As String, s_Trabajador As String
  Dim s_Parametro(10) As String, n_Importe As Double
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim nDiasMes As Integer, nDias As Integer
  
  ' Recupero los parametros de remuneraciones
  s_Sql = "SELECT cpcremuies, cpcremuonp, cpcremuessalud, cpcremuartista, cpcremuquinta, "
  s_Sql = s_Sql & "cpcies, cpconp, cpcessalud, cpcartista, cpcquinta "
  s_Sql = s_Sql & "FROM plparametroafp "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
  
  ' Obtengo los conceptos de recuperación
  a_Parametro(1) = gdl_Funcion.aTexto(porstRecordset!cpcremuies)
  a_Parametro(2) = gdl_Funcion.aTexto(porstRecordset!cpcremuonp)
  a_Parametro(3) = gdl_Funcion.aTexto(porstRecordset!cpcremuessalud)
  a_Parametro(4) = gdl_Funcion.aTexto(porstRecordset!cpcremuartista)
  a_Parametro(5) = gdl_Funcion.aTexto(porstRecordset!cpcremuquinta)
  a_Parametro(6) = gdl_Funcion.aTexto(porstRecordset!cpcies)
  a_Parametro(7) = gdl_Funcion.aTexto(porstRecordset!cpconp)
  a_Parametro(8) = gdl_Funcion.aTexto(porstRecordset!cpcessalud)
  a_Parametro(9) = gdl_Funcion.aTexto(porstRecordset!cpcartista)
  a_Parametro(10) = gdl_Funcion.aTexto(porstRecordset!cpcquinta)
  porstRecordset.Close
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Procesando Información ..."
  
  ' Recupero la información para exportar
  s_Sql = "SELECT DISTINCTROW res.codcls, res.pdoano, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
  s_Sql = s_Sql & "psn.coddci, psn.numdociden, dxr.regpension, res.codcpc, SUM(IFNULL(asi.diatrabajo, 0)) AS dias, SUM(IFNULL(asi.diavacaciones, 0)) AS diavaca, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0)), 2) AS remuneracion "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "WHERE res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.pdomes='" & s_Periodo & "' "
  s_Sql = s_Sql & "AND res.codcls IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.codcpc IN ('" & a_Parametro(1) & "', '" & a_Parametro(2) & "', '" & a_Parametro(3) & "', '" & a_Parametro(4) & "', '" & a_Parametro(5) & "', "
  s_Sql = s_Sql & "'" & a_Parametro(6) & "', '" & a_Parametro(7) & "', '" & a_Parametro(8) & "', '" & a_Parametro(9) & "', '" & a_Parametro(10) & "') "
  s_Sql = s_Sql & "GROUP BY res.codcls, res.codpsn, res.codcpc "
  s_Sql = s_Sql & "ORDER BY res.codcls, res.codpsn, res.codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    nDiasMes = (gdl_Funcion.NumeroDiasMes(s_Periodo, ps_Anyo) - 27)
    nDiasMes = Choose(nDiasMes, -2, -1, 0, 1)
    
    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("codcls", "codpsn", "nompsn", "coddci", "numdociden", "dias", "remies", "remonp", "remessalud", "remartista", "remquinta", "quinta")
      a_Valores = Array("", "", "", "", "", CInt(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0))
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
      s_Caracter = "|"
    End If
    While Not porstRecordset.EOF
      nDias = CInt(porstRecordset!dias) + CInt(porstRecordset!diavaca)
      nDias = nDias + IIf(nDias = 0, nDias, nDiasMes)
      ' Genero el registro de grabación
      psRegistro = ""
      psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!coddci) & s_Caracter
      psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!numdociden) & s_Caracter
      psRegistro = psRegistro & nDias & s_Caracter
      s_Trabajador = porstRecordset!codpsn
      s_Parametro(1) = "": s_Parametro(2) = "": s_Parametro(3) = "": s_Parametro(4) = ""
      s_Parametro(5) = "": s_Parametro(6) = "": s_Parametro(7) = "": s_Parametro(8) = ""
      s_Parametro(9) = "": s_Parametro(10) = ""
      If s_Accion = "R" Then
        a_Valores(0) = porstRecordset("codcls"): a_Valores(1) = porstRecordset("codpsn")
        a_Valores(2) = porstRecordset("nompsn"): a_Valores(3) = porstRecordset("coddci")
        a_Valores(4) = porstRecordset("numdociden"): a_Valores(5) = nDias
      End If
      Do
        n_Importe = CDec(porstRecordset("remuneracion"))
        ' Remuneración de IES
        If a_Parametro(1) = porstRecordset("codcpc") And porstRecordset("regpension") = "4" And n_Importe > 0 Then
          s_Parametro(1) = Format(n_Importe, "###########0.00")
        End If
        ' Remuneración de ONP
        If a_Parametro(2) = porstRecordset("codcpc") And porstRecordset("regpension") = "1" And n_Importe > 0 Then
          s_Parametro(2) = Format(n_Importe, "###########0.00")
        End If
        ' Remuneración de salud(essalud)
        If a_Parametro(3) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(3) = Format(n_Importe, "###########0.00")
        End If
        ' Remuneración de fondo del artista
        If a_Parametro(4) = porstRecordset("codcpc") And porstRecordset("regpension") = "5" And n_Importe > 0 Then
          s_Parametro(4) = Format(n_Importe, "###########0.00")
        End If
        ' Remuneración de quinta
        If a_Parametro(5) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(5) = Format(n_Importe, "###########0.00")
        End If
        ' Descuento de IES
        If a_Parametro(6) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(6) = Format(n_Importe, "###########0.00")
        End If
        ' Descuento de ONP
        If a_Parametro(7) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(7) = Format(n_Importe, "###########0.00")
        End If
        ' Aporte de salud(essalud)
        If a_Parametro(8) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(8) = Format(n_Importe, "###########0.00")
        End If
        ' Aporte de fondo del artista
        If a_Parametro(9) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(9) = Format(n_Importe, "###########0.00")
        End If
        ' Descuento de quinta
        If a_Parametro(10) = porstRecordset("codcpc") And n_Importe > 0 Then
          s_Parametro(10) = Format(n_Importe, "###########0.00")
        End If
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        DoEvents
        porstRecordset.MoveNext
        ' Fin de archivo
        If porstRecordset.EOF Then Exit Do
      Loop While s_Trabajador = porstRecordset!codpsn
      
      ' Validacion de remuneracion con descunto o aporte afectos
      For n_Index = 1 To 5
        s_Parametro(n_Index) = IIf(s_Parametro(n_Index + 5) = "", "", s_Parametro(n_Index))
      Next n_Index
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores(6) = CDec(IIf(s_Parametro(1) = "", 0, s_Parametro(1)))
        a_Valores(7) = CDec(IIf(s_Parametro(2) = "", 0, s_Parametro(2)))
        a_Valores(8) = CDec(IIf(s_Parametro(3) = "", 0, s_Parametro(3)))
        a_Valores(9) = CDec(IIf(s_Parametro(4) = "", 0, s_Parametro(4)))
        a_Valores(10) = CDec(IIf(s_Parametro(5) = "", 0, s_Parametro(5)))
        a_Valores(11) = CDec(IIf(s_Parametro(10) = "", 0, s_Parametro(10)))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = psRegistro & s_Parametro(1) & s_Caracter
        psRegistro = psRegistro & s_Parametro(2) & s_Caracter
        psRegistro = psRegistro & s_Parametro(3) & s_Caracter
        psRegistro = psRegistro & s_Parametro(4) & s_Caracter
        psRegistro = psRegistro & s_Parametro(5) & s_Caracter
        psRegistro = psRegistro & s_Parametro(10) & s_Caracter
        potxtFileExp.WriteLine psRegistro
      End If
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
Private Sub ExportaTrabajadores(ByVal s_Archivo As String, ByVal s_Periodo As String, s_Proceso As String, s_FechaHora As String, s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim s_CodEps As String, s_Estado As String
  
  ' Recupero la información para exportar
  s_Sql = "SELECT psn.codcls, psn.codpsn, psn.coddci, dci.desdci, psn.numdociden, psn.apepaterno, psn.apematerno, "
  s_Sql = s_Sql & "psn.nombres, psn.fecnacimiento, psn.sexopsn, psn.telefono, psn.fecingreso, "
  s_Sql = s_Sql & "psn.estadopsn, psn.codtpt, tpt.destpt, psn.fecbaja, psn.codeps, eps.deseps, "
  s_Sql = s_Sql & "eps.ruceps, psn.essvida, psn.regpension, psn.cobsctr, psn.fecingregpen, "
  s_Sql = s_Sql & "psn.nomviadirec, psn.numerdirec, psn.intedirec, psn.nomzondirec, "
  s_Sql = s_Sql & "psn.refedirec, psn.codvia, via.desvia, psn.codzona, zon.deszona, psn.ubigeodir "
  s_Sql = s_Sql & "FROM plpersonal psn "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN pltpotrabajador tpt ON psn.codtpt=tpt.codtpt "
  s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
  s_Sql = s_Sql & "WHERE psn.codcls IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND (DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & s_Periodo & "' AND psn.estadopsn<>'I') "
  s_Sql = s_Sql & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & s_Periodo & "' AND psn.estadopsn='I') "
  s_Sql = s_Sql & "ORDER BY codcls, codpsn"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando información ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    
    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("codcls", "codpsn", "coddci", "numdociden", "apepaterno", "apematerno", "nombres", "fecnacimiento", "sexopsn", "telefono", "fecingreso", "estadopsn", "codtpt", "fecbaja", "ruceps", "essvida", "regpension", "cobsctr", "fecingregpen", "nomviadirec", "numerdirec", "intedirec", "nomzondirec", "refedirec", "codvia", "codzona", "ubigeodir")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
      s_Caracter = "|"
    End If
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      s_CodEps = gdl_Funcion.aTexto(porstRecordset!codeps)
      s_Estado = gdl_Funcion.aTexto(porstRecordset!estadopsn)
      If s_Accion = "R" Then
        If s_CodEps = "I" Then
          s_Estado = IIf(s_CodEps = "" Or s_CodEps = "99", "13", "12")
        ElseIf s_CodEps = "L" Then
          s_Estado = IIf(s_CodEps = "" Or s_CodEps = "99", "15", "14")
        Else
          s_Estado = IIf(s_CodEps = "" Or s_CodEps = "99", "11", "10")
        End If
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        a_Valores = Array(porstRecordset!codcls, porstRecordset!codpsn, gdl_Funcion.aTexto(porstRecordset!coddci), gdl_Funcion.aTexto(porstRecordset!numdociden), gdl_Funcion.aTexto(porstRecordset!apepaterno), gdl_Funcion.aTexto(porstRecordset!apematerno), gdl_Funcion.aTexto(porstRecordset!nombres), Format(porstRecordset!fecnacimiento, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!sexopsn), gdl_Funcion.aTexto(porstRecordset!telefono), _
                    Format(porstRecordset!fecingreso, s_FmtFechMysql_0), s_Estado, gdl_Funcion.aTexto(porstRecordset!codtpt), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!ruceps), gdl_Funcion.aTexto(porstRecordset!essvida), gdl_Funcion.aTexto(porstRecordset!regpension), gdl_Funcion.aTexto(porstRecordset!cobsctr), Format(porstRecordset!fecingregpen, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!nomviadirec), _
                    gdl_Funcion.aTexto(porstRecordset!numerdirec), gdl_Funcion.aTexto(porstRecordset!intedirec), gdl_Funcion.aTexto(porstRecordset!nomzondirec), gdl_Funcion.aTexto(porstRecordset!refedirec), gdl_Funcion.aTexto(porstRecordset!codvia), gdl_Funcion.aTexto(porstRecordset!codzona), gdl_Funcion.aTexto(porstRecordset!ubigeodir))
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        psRegistro = ""
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!coddci) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!numdociden) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!apepaterno) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!apematerno) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!nombres) & s_Caracter
        psRegistro = psRegistro & Format(porstRecordset!fecnacimiento, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & Val(gdl_Funcion.aTexto(porstRecordset!sexopsn)) + 1 & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!telefono) & s_Caracter
        psRegistro = psRegistro & Format(porstRecordset!fecingreso, s_FormatoFecha) & s_Caracter
        If s_CodEps = "I" Then
          psRegistro = psRegistro & IIf(s_CodEps = "" Or s_CodEps = "99", "13", "12") & s_Caracter
        ElseIf s_CodEps = "L" Then
          psRegistro = psRegistro & IIf(s_CodEps = "" Or s_CodEps = "99", "15", "14") & s_Caracter
        Else
          psRegistro = psRegistro & IIf(s_CodEps = "" Or s_CodEps = "99", "11", "10") & s_Caracter
        End If
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codtpt) & s_Caracter
        psRegistro = psRegistro & Format(porstRecordset!fecbaja, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!ruceps) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!essvida) & s_Caracter
        psRegistro = psRegistro & Val(gdl_Funcion.aTexto(porstRecordset!regpension)) + 1 & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!cobsctr) & s_Caracter
        psRegistro = psRegistro & Format(porstRecordset!fecingregpen, s_FormatoFecha) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!nomviadirec) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!numerdirec) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!intedirec) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!nomzondirec) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!refedirec) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codvia) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!codzona) & s_Caracter
        psRegistro = psRegistro & gdl_Funcion.aTexto(porstRecordset!ubigeodir) & s_Caracter
        potxtFileExp.WriteLine psRegistro
      End If
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
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
  Dim s_Copias As String, s_OldMessage As String
  Dim s_Archivo As String
  Dim s_Extension As String, s_Descripcion As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  Select Case Index
   Case 0  ' Actualización de parametros
    If cmbPeriodo = "" Then Beep: MsgBox "Debe selecionar el Periodo de Información", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    fPrmExpInformacion.Show vbModal
   Case 1     ' Genera el archivo de transferencia de información
    ' Verifico que existan registros seleccionados
    If cmbPeriodo = "" Then Beep: MsgBox "Debe selecionar el Periodo de Información", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    
    ' Verifico que existan parametros de exportación
    If ribParametro(0).Value Then
      s_Sql = "SELECT cpcremuies, cpcremuonp, cpcremuessalud, "
      s_Sql = s_Sql & "cpcremuartista , cpcremuquinta, cpcquinta "
      s_Sql = s_Sql & "FROM plparametroafp "
      s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
      If gdl_Funcion.aTexto(porstRecordset!cpcremuessalud) = "" Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
    End If
    
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Proceso = "rptexpinfo"
    s_Extension = IIf(ribParametro(0).Value, ".djt", IIf(ribParametro(1).Value, ".ase", ".der"))
    s_Archivo = IIf(ribParametro(0).Value, "0600" & ps_Anyo & Left(cmbPeriodo, 2), "") & ps_RucEmpresa & s_Extension
    s_Descripcion = IIf(ribParametro(0).Value, "Archivos de remuneraciones(*.djt)|*.djt", IIf(ribParametro(1).Value, "Archivos de asegurados(*.ase)|*.ase", "Archivos de derechohabientes(*.der)|*.der")) & "|Todos los archivos(*.*)|*.*"
    
    On Error GoTo CancelaDialogo
    fMenu.cdlDialogo.DialogTitle = "Grabar Archivo Como"
    fMenu.cdlDialogo.CancelError = True
    fMenu.cdlDialogo.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    fMenu.cdlDialogo.FileName = s_Archivo
    fMenu.cdlDialogo.DefaultExt = s_Extension
    fMenu.cdlDialogo.Filter = s_Descripcion
    fMenu.cdlDialogo.ShowSave
  
CancelaDialogo:
    ' verifico si existe error y desactivo
    If Not Err.Number = 0 Then
      MsgBox Error(Err.Number)
      Exit Sub
    End If
    On Error GoTo 0
    
    ChDir App.path
    If MsgBox("¿ Estás Seguro de Generar Archivo para el PDT - " & IIf(ribParametro(0).Value, "Remuneraciones", IIf(ribParametro(1).Value, "Asegurados", "Derechohabientes")) & "? ", vbQuestion + vbYesNo) = vbYes Then
      s_Archivo = fMenu.cdlDialogo.FileName
      ' Barro el arreglo de registros marcadas (bookmarks)
      For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
        tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
        gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      Next n_Index
      If ribParametro(0).Value Then
        ExportaRemuneraciones s_Archivo, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "G"
      ElseIf ribParametro(1).Value Then
        ExportaTrabajadores s_Archivo, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "G"
      Else
        ExportaDerechoHabiente s_Archivo, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "G"
      End If
      MsgBox "Proceso de Exportación Finalizo con Exito", vbInformation
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, "", ps_Usuario, s_FechaHora, "E"
    End If
    
    ChDrive Left$(App.path, 1)
    ChDir App.path
    
   Case 2 ' Busqueda de registro
    Set go_tdbBusqueda = tdbRegistro
    Set go_dcaBusqueda = dcaRegistro
    gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
    fBusqueda.Show vbModal
   Case 3, 4, 5 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), (Index - 4), s_TitleTable
   Case 6, 7    ' Opciones de impresión
    ' Verifico que existan registros seleccionados
    If cmbPeriodo = "" Then Beep: MsgBox "Debe selecionar el Periodo de Información", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    ' Verifico que existan parametros de exportación
    If ribParametro(0).Value Then
      s_Sql = "SELECT cpcremuies, cpcremuonp, cpcremuessalud, cpcremuartista, cpcremuquinta, "
      s_Sql = s_Sql & "cpcies, cpconp, cpcessalud, cpcartista, cpcquinta "
      s_Sql = s_Sql & "FROM plparametroafp "
      s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
      If gdl_Funcion.aTexto(porstRecordset!cpcremuessalud) = "" Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
    End If
    
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Proceso = "rptexpinfo"
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = Me.Caption
    gdl_Procedure.ps_ReportName = IIf(ribParametro(0).Value, "rptremupdt", IIf(ribParametro(1).Value, "rptasegpdt", "rptdehapdt"))
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del store procedure
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = "": aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = ""
    aElemento(1, 2) = ""
    ' Campos de Parametros del Reporte
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & UCase(gdl_Procedure.ps_ReportTitle) & IIf(ribParametro(0).Value, " - REMUNERACIONES", IIf(ribParametro(1).Value, " - TRABAJADORES", " - DERECHOHABIENTES")) & ";true"
    aElemento(2, 2) = "Periodo;" & cmbPeriodo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = ""
    aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If ribParametro(0).Value Then
      s_Sql = s_Sql & "codcls char(2) Not Null, codpsn varchar(11) Not Null, "
      s_Sql = s_Sql & "nompsn varchar(80) Null, coddci char(2) Null, numdociden varchar(11) Null, "
      s_Sql = s_Sql & "dias smallint(3) Null Default '0', remies decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "remonp decimal(18,2) Null Default '0', remessalud decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "remartista decimal(18,2) Null Default '0', remquinta decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "quinta decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "PRIMARY KEY (codcls, codpsn)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' Archivo temporal de impresión
      ExportaRemuneraciones "tmp" & gdl_Procedure.ps_ReportName, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "R"
    ElseIf ribParametro(1).Value Then
      s_Sql = s_Sql & "codcls char(2) Not Null, codpsn varchar(11) Not Null, coddci char(2) Null, "
      s_Sql = s_Sql & "numdociden varchar(11) Null, apepaterno varchar(25) Null, "
      s_Sql = s_Sql & "apematerno varchar(25) Null, nombres varchar(25) Null, "
      s_Sql = s_Sql & "fecnacimiento date Null, sexopsn char(1) Null, "
      s_Sql = s_Sql & "telefono varchar(10) Null, fecingreso date Null, "
      s_Sql = s_Sql & "estadopsn char(2) Null, codtpt char(2) Null, "
      s_Sql = s_Sql & "fecbaja date Null, ruceps varchar(11) Null, "
      s_Sql = s_Sql & "essvida char(1) Null, regpension char(11) Null, "
      s_Sql = s_Sql & "cobsctr char(1) Null, fecingregpen date Null, "
      s_Sql = s_Sql & "nomviadirec varchar(20) Null, numerdirec varchar(4) Null, "
      s_Sql = s_Sql & "intedirec varchar(4) Null, nomzondirec varchar(20) Null, "
      s_Sql = s_Sql & "refedirec varchar(50) Null, codvia char(2) Null, "
      s_Sql = s_Sql & "codzona char(2) Null, ubigeodir varchar(6) Null, "
      s_Sql = s_Sql & "PRIMARY KEY (codcls, codpsn)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' Archivo temporal de impresión
      ExportaTrabajadores "tmp" & gdl_Procedure.ps_ReportName, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "R"
    Else
      s_Sql = s_Sql & "codcls char(2) Not Null, codpsn varchar(11) Not Null, orden smallint(2) Not Null, "
      s_Sql = s_Sql & "coddci char(2) Null, numdociden varchar(11) Null, "
      s_Sql = s_Sql & "apepaterno varchar(25) Null, apematerno varchar(25) Null, "
      s_Sql = s_Sql & "nombres varchar(25) Null, coddcifam char(2) Null, "
      s_Sql = s_Sql & "numdocidenfam varchar(11) Null, apepaternofam varchar(25) Null, "
      s_Sql = s_Sql & "apematernofam varchar(25) Null, nombresfam varchar(25) Null, "
      s_Sql = s_Sql & "fecnacimientofam date Null, sexofam char(1) Null, "
      s_Sql = s_Sql & "vinculo char(1) Null, caratamed varchar(20) Null, "
      s_Sql = s_Sql & "estadofam char(2) Null, motivoina char(1) Null, "
      s_Sql = s_Sql & "certificadomed varchar(20) Null, domicilio char(1) Null, "
      s_Sql = s_Sql & "nomviadom varchar(20) Null, numerdom varchar(4) Null, "
      s_Sql = s_Sql & "intedom varchar(4) Null, nomzonadom varchar(20) Null, "
      s_Sql = s_Sql & "refedom varchar(50) Null, codvia char(2) Null, "
      s_Sql = s_Sql & "codzona char(2) Null, ubigeodom varchar(6) Null, "
      s_Sql = s_Sql & "PRIMARY KEY (codcls, codpsn, orden)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' Archivo temporal de impresión
      ExportaDerechoHabiente "tmp" & gdl_Procedure.ps_ReportName, Left(cmbPeriodo, 2), s_Proceso, s_FechaHora, "R"
    End If
    ' Genera la información del reporte
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, "", ps_Usuario, s_FechaHora, "E"
  End Select

End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6060: Me.Width = 6230
  Me.Left = 520: Me.Top = 300
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Exportación de Información a la SUNAT"
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
  ReDim aElemento(8, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "promedio", "genarchi", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Parametros", "Generación de Archivo", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
 '[ Configuración el control de ayuda
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ' Tipo de analisis
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Información " & Choose(n_Index + 1, "Remuneraciones", "Datos del Personal", "Datos Derechohabientes")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "apciano", "asiperso", "familiar") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribParametro(0).Value = True
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
