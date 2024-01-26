VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fTablaSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "tabsiste.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   6135
   Begin Threed.SSFrame frmCuadro 
      Height          =   675
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   5250
      _Version        =   65536
      _ExtentX        =   9260
      _ExtentY        =   1191
      _StockProps     =   14
      Caption         =   " Tablas Sistema "
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
      ShadowStyle     =   1
      Begin VB.ComboBox cmbTabla 
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
         Height          =   315
         Left            =   1425
         TabIndex        =   2
         Top             =   255
         Width           =   3525
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro :"
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
         Left            =   240
         TabIndex        =   1
         Top             =   285
         Width           =   900
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4605
      Left            =   45
      TabIndex        =   13
      Top             =   795
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
      Top             =   5445
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
      Height          =   5655
      Index           =   0
      Left            =   5370
      TabIndex        =   3
      Top             =   120
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9975
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
         TabIndex        =   5
         Tag             =   "0"
         Top             =   1095
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1515
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   14737632
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   2355
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2790
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   4020
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   4455
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":0098
      End
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
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Tag             =   "0"
         Top             =   660
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   3210
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "tabsiste.frx":00D0
      End
   End
End
Attribute VB_Name = "fTablaSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, i As Integer            ' Indice de la barra de herramientas, indice para bucle
Private s_swEntidad As String                           ' Swicht de tabla de sistema
Private Sub RecuperaRegistros(ByVal s_Orden As String)
  Dim s_Order As String, s_BaseDatos As String
  
  ' Campos de las columnas
  For i = 0 To 2
    If s_swEntidad = "cta" Then
      tdbRegistro.Columns(i).DataField = Choose(i + 1, "cod", "det", "est") & s_swEntidad
    Else
      tdbRegistro.Columns(i).DataField = Choose(i + 1, "cod", "des", "estado") & s_swEntidad
    End If
  Next i
  tdbRegistro.Columns(0).Caption = "Código"
  tdbRegistro.Columns(0).HeadAlignment = vbLeftJustify
  tdbRegistro.Columns(0).Alignment = vbLeftJustify
  tdbRegistro.Columns(0).NumberFormat = ""
  
  s_Order = ""
  s_BaseDatos = ps_DataBase
  ' Cadenas de Texto, Recuperar Información
  Select Case s_swEntidad
   Case "act"
    s_Sql = "SELECT codact, desact, estadoact "
    s_Sql = s_Sql & "FROM plactividad "
   Case "cdt"
    s_Sql = "SELECT codcdt, descdt, estadocdt "
    s_Sql = s_Sql & "FROM plconditrabajo "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
   Case "cgo"
    s_Sql = "SELECT codcgo, descgo, estadocgo "
    s_Sql = s_Sql & "FROM plcargo "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
   Case "cls"
    s_Sql = "SELECT codcls, descls, clave, horadiaria, fmtboleta, estadocls, tipo "
    s_Sql = s_Sql & "FROM plclasplan "
   Case "cpc"
    s_Sql = "SELECT codcpc, descpc, aliascpc, tipocpc, obs,estadocpc "
    s_Sql = s_Sql & "FROM plconcepto "
   Case "cta"
    s_Sql = "SELECT codcta, detcta, tpocta, natcta, tposdo, inddoc, indcco, (CASE WHEN estcta='A' THEN 1 ELSE 0 END) AS estcta "
    s_Sql = s_Sql & "FROM cocta "
    s_BaseDatos = ps_DaBasCon
   Case "dmo"
    s_Sql = "SELECT CONCAT(codmon, valordmo) AS cPrimaryKey, codmon, valordmo, desdmo, estadodmo "
    s_Sql = s_Sql & "FROM pldstmoneda "
    s_Order = "codmon, "
    ' Formato de columnas
    tdbRegistro.Columns(0).DataField = "valor" & s_swEntidad
    tdbRegistro.Columns(0).Caption = "Valor"
    tdbRegistro.Columns(0).HeadAlignment = vbRightJustify
    tdbRegistro.Columns(0).Alignment = vbRightJustify
    tdbRegistro.Columns(0).NumberFormat = "standard"
   Case "dci"
    s_Sql = "SELECT coddci, desdci, sigladci, codsunat, estadodci "
    s_Sql = s_Sql & "FROM pldocidentidad "
   Case "bco"
    s_Sql = "SELECT codbco, desbco, cuentamn, cuentame, codentidad, formato, impolimite_mn, impolimite_me, estadobco "
    s_Sql = s_Sql & "FROM plbanco "
   Case "eps"
    s_Sql = "SELECT codeps, deseps, ruceps, factoreps, codsunat, estadoeps "
    s_Sql = s_Sql & "FROM plentidadeps "
   Case "ldn"
    s_Sql = "SELECT codldn, desldn, estadoldn "
    s_Sql = s_Sql & "FROM plcodigoldn "
   Case "mfo"
    s_Sql = "SELECT codmfo, desmfo, estadomfo "
    s_Sql = s_Sql & "FROM plmodforma "
   Case "bdh"
    s_Sql = "SELECT codbdh, desbdh, estadobdh "
    s_Sql = s_Sql & "FROM plmotbajadh "
   Case "mof"
    s_Sql = "SELECT codmof, desmof, estadomof "
    s_Sql = s_Sql & "FROM plmotfin "
   Case "nac"
    s_Sql = "SELECT codnac, desnac, codpemi, estadonac "
    s_Sql = s_Sql & "FROM plnacionalidad "
   Case "niv"
    s_Sql = "SELECT codniv, desniv, estadoniv "
    s_Sql = s_Sql & "FROM plniveducativo "
   Case "prd"
    s_Sql = "SELECT codprd, desprd, estadoprd "
    s_Sql = s_Sql & "FROM plperiodicidad "
   Case "pfs"
    s_Sql = "SELECT codpfs, despfs, estadopfs "
    s_Sql = s_Sql & "FROM plprofesion "
   Case "sec"
    s_Sql = "SELECT codsec, dessec, codintersec, estadosec "
    s_Sql = s_Sql & "FROM plseccion "
   Case "stp"
    s_Sql = "SELECT codstp, desstp, estadostp "
    s_Sql = s_Sql & "FROM plsitrapen "
   Case "tic"
    s_Sql = "SELECT codtic, destic, estadotic "
    s_Sql = s_Sql & "FROM pltipcom "
   Case "tco"
    s_Sql = "SELECT codtco, destco, estadotco "
    s_Sql = s_Sql & "FROM pltipcontrato "
   Case "est"
    s_Sql = "SELECT codest, desest, estadoest "
    s_Sql = s_Sql & "FROM plestablecimiento "
   Case "tip"
    s_Sql = "SELECT codtip, destip, estadotip "
    s_Sql = s_Sql & "FROM pltippago "
   Case "tsu"
    s_Sql = "SELECT codtsu, destsu, estadotsu "
    s_Sql = s_Sql & "FROM pltipsusp "
   Case "via"
    s_Sql = "SELECT codvia, desvia, abrevia, estadovia "
    s_Sql = s_Sql & "FROM pltipovia "
   Case "zona"
    s_Sql = "SELECT codzona, deszona, abrezona, estadozona "
    s_Sql = s_Sql & "FROM pltipozona "
   Case "tpt"
    s_Sql = "SELECT codtpt, destpt, estadotpt "
    s_Sql = s_Sql & "FROM pltpotrabajador "
   Case "ubica"
    s_Sql = "SELECT codubica, desubica, codinterubica, estadoubica "
    s_Sql = s_Sql & "FROM plubicacion "
   Case "vfa"
    s_Sql = "SELECT codvfa, desvfa, estadovfa "
    s_Sql = s_Sql & "FROM plvinfami "
   Case "con"
    s_Sql = "SELECT codcon, descon, estadocon, tipcon "
    s_Sql = s_Sql & "FROM plconcesunat "
   Case "cao"
    s_Sql = "SELECT codcao, descao, estadocao "
    s_Sql = s_Sql & "FROM plcatocu "
   Case "ctr"
    s_Sql = "SELECT codctr, desctr, estadoctr "
    s_Sql = s_Sql & "FROM plconven "
  End Select
  s_Orden = s_Order & IIf(s_Orden = "", tdbRegistro.Columns(0).DataField & " ASC", s_Orden)
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & s_BaseDatos, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly
  
End Sub
Private Sub cmbTabla_Click()
  
  ' Especificaciones adicionales
  s_swEntidad = Choose(cmbTabla.ListIndex + 1, "act", "cgo", "cls", "cpc", "cdt", "cta", "dmo", "dci", "bco", "eps", "ldn", "mfo", "bdh", "mof", "nac", "niv", "prd", "pfs", "sec", "stp", "tic", "tco", "est", "tip", "tsu", "via", "zona", "tpt", "ubica", "vfa", "con", "cao", "ctr")
  s_TitleTable = cmbTabla.Text
  tdbRegistro.Caption = s_TitleTable
  RecuperaRegistros ""
  
End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_Orden As String, s_BaseDatos As String

  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0, 1, 2 ' Visualizar o analizar, nuevo, eliminar registro
   ' If (Index = 1 And ps_DaBasCon = ps_DataBase) Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: tasRegister.Tabs(1).Selected = True: txtCodigo.SetFocus: Exit Sub
    If (Index <> 1 And (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF)) Then Exit Sub
    Me.Tag = Choose(Index + 1, s_MdoData_Vis, s_MdoData_Ins, s_MdoData_Del)
    ' Accion de acuerdo a la tabla
    Select Case s_swEntidad
     Case "act": fAbcActividad.Show
     Case "bdh": fAbcMotBajaDH.Show
     Case "cao": fabcCatOcu.Show
     Case "con": fAbcSunat.Show
     Case "ctr": fabcConven.Show
     Case "ldn": fAbcLargaDistancia.Show
     Case "mfo": fAbcModForma.Show
     Case "mof": fAbcMotFin.Show
     Case "sit": fAbcSitrapen.Show
     Case "tic": fAbcTipComp.Show
     Case "tip": fAbcTipPago.Show
     Case "tsu": fAbcTipSusp.Show
     Case "vfa": fAbcVincFami.Show
     Case "niv": fAbcNivEducativo.Show
     Case "bco": fAbcEntidadBanco.Show
     Case "cdt": fAbcCondicionTrabajo.Show
     Case "cgo": fAbcCargoPersonal.Show
     Case "cls": fAbcClasePlanilla.Show
     Case "cpc": fAbcConcepCalculo.Show
     Case "cta": fAbcCuenta.Show
     Case "dci": fAbcDocIdentidad.Show
     Case "dmo": fAbcDistribuMoneda.Show
     Case "eps": fAbcEntidadEps.Show
     Case "est": fAbcEstablecimiento.Show
     Case "nac": fAbcNacionalidad.Show
     Case "pfs": fAbcProfesion.Show
     Case "prd": fAbcPeriodicidad.Show
     Case "sec": fAbcSeccion.Show
     Case "stp": fAbcSitrapen.Show
     Case "tco": fAbcTipContrato.Show
     Case "tpt": fAbcTipoTrabajador.Show
     Case "ubica": fAbcUbicacion.Show
     Case "via": fAbcTipoVia.Show
     Case "zona": fAbcTipoZona.Show
    End Select
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
    
    ' Parametros de Impresión
    ReDim aElemento(3, 2): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    gdl_Procedure.ps_ReportTitle = UCase(tdbRegistro.Caption)
    s_Orden = ""
    s_BaseDatos = ps_DataBase
    ' Accion de acuerdo a la tabla
    Select Case s_swEntidad
     Case "act"
      gdl_Procedure.ps_ReportName = "lstactividad"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codact, desact, estadoact "
      s_Sql = s_Sql & "FROM plactividad "
     Case "bco"
      gdl_Procedure.ps_ReportName = "lstentidadbanco"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codbco, desbco, cuentamn, cuentame, codentidad, formato, estadobco "
      s_Sql = s_Sql & "FROM plbanco "
     Case "bdh"
      gdl_Procedure.ps_ReportName = "lstmotbajadh"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codbdh, desbdh, estadobdh "
      s_Sql = s_Sql & "FROM plmotbajadh "
     Case "cao"
      gdl_Procedure.ps_ReportName = "lstcatocu"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codcao, descao "
      s_Sql = s_Sql & "FROM plcatocu "
     Case "cdt"           ' condición de trabajo
      gdl_Procedure.ps_ReportName = "lstcargo"
      s_Sql = "SELECT codcdt, descdt, estadocdt "
      s_Sql = s_Sql & "FROM plconditrabajo "
      s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "' "
     Case "cgo"
      gdl_Procedure.ps_ReportName = "lstcargo"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codcgo, descgo, estadocgo "
      s_Sql = s_Sql & "FROM plcargo "
      s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "' "
     Case "cls"
      gdl_Procedure.ps_ReportName = "lstclsplanilla"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codcls, descls, clave, estadocls "
      s_Sql = s_Sql & "FROM plclasplan "
     Case "con"
      gdl_Procedure.ps_ReportName = "lstconsun"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codcon, descon "
      s_Sql = s_Sql & "FROM plconcesunat "
     Case "cpc"
      gdl_Procedure.ps_ReportName = "lstconcecalculo"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codcpc, descpc, aliascpc, tipocpc, estadocpc,obs "
      s_Sql = s_Sql & "FROM plconcepto "
     Case "cta"
       gdl_Procedure.ps_ReportName = "lstctacontable"
      'Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codcta, detcta, "
      s_Sql = s_Sql & "(CASE TpoCta WHEN " & s_Estado_Ina & " THEN 'Título' WHEN " & s_Estado_Act & " THEN 'Detalle' END) AS tpocta, "
      s_Sql = s_Sql & "(CASE NatCta WHEN " & s_Estado_Ina & " THEN 'Deudora' WHEN " & s_Estado_Act & " THEN 'Acreedora' END) AS natcta, "
      s_Sql = s_Sql & "(CASE TpoSdo "
      s_Sql = s_Sql & "WHEN 'I' THEN 'Inv.' "
      s_Sql = s_Sql & "WHEN 'R' THEN 'Res.' "
      s_Sql = s_Sql & "WHEN 'F' THEN 'Func.' "
      s_Sql = s_Sql & "WHEN 'N' THEN 'Nat.' "
      s_Sql = s_Sql & "WHEN 'A' THEN 'F/N' END) AS tposdo, "
      s_Sql = s_Sql & "(CASE IndDoc WHEN " & s_Estado_Ina & " THEN 'No' WHEN " & s_Estado_Act & " THEN 'Si' END) AS indaux, "
      s_Sql = s_Sql & "(CASE IndCCo WHEN " & s_Estado_Ina & " THEN 'No' WHEN " & s_Estado_Act & " THEN 'Si' END) AS indcco, "
      s_Sql = s_Sql & "(CASE TpoMon WHEN 'N' THEN 'MN'  WHEN 'E' THEN 'ME' END) AS tpomon, "
      s_Sql = s_Sql & "(CASE EstCta WHEN 'A' THEN 'Activa' WHEN 'I' THEN 'Inactiva' END) AS estcta "
      s_Sql = s_Sql & "FROM cocta "
      s_Sql = s_Sql & "ORDER BY codcta"
      s_BaseDatos = ps_DaBasCon
     Case "ctr"
      gdl_Procedure.ps_ReportName = "lstconven"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codctr, desctr "
      s_Sql = s_Sql & "FROM plconven "
     Case "dci"
      gdl_Procedure.ps_ReportName = "lstdocuidentid"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT coddci, desdci, sigladci,codsunat, estadodci "
      s_Sql = s_Sql & "FROM pldocidentidad "
     Case "dmo"
      gdl_Procedure.ps_ReportName = "lstdismonedas"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT CONCAT(codmon, valordmo) AS cPrimaryKey, codmon, valordmo, desdmo, estadodmo, "
      s_Sql = s_Sql & "(CASE WHEN codmon='N' THEN '" & s_Codmon_mn_Txt & "' ELSE '" & s_Codmon_me_Txt & "' END) AS simbolomon "
      s_Sql = s_Sql & "FROM pldstmoneda "
      s_Orden = "codmon, "
     Case "eps"
      gdl_Procedure.ps_ReportName = "lstentidadeps"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codeps, deseps, ruceps, factoreps, estadoeps, codsunat "
      s_Sql = s_Sql & "FROM plentidadeps "
     Case "est"
      gdl_Procedure.ps_ReportName = "lstestablecimiento"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codest, desest, estadoest "
      s_Sql = s_Sql & "FROM plestablecimiento "
     Case "ldn"
      gdl_Procedure.ps_ReportName = "lstlargadistancia"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codldn, desldn, estadoldn "
      s_Sql = s_Sql & "FROM plcodigoldn "
     Case "mfo"
      gdl_Procedure.ps_ReportName = "lstmodforma"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codmfo, desmfo, estadomfo "
      s_Sql = s_Sql & "FROM plmodforma "
     Case "mof"
      gdl_Procedure.ps_ReportName = "lstmotfin"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codmof, desmof, estadomof "
      s_Sql = s_Sql & "FROM plmotfin "
     Case "nac"
      gdl_Procedure.ps_ReportName = "lstnacionalidad"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codnac, desnac, estadonac "
      s_Sql = s_Sql & "FROM plnacionalidad "
     Case "nac"
      gdl_Procedure.ps_ReportName = "lstniveducativo"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codniv, desniv, estadoniv "
      s_Sql = s_Sql & "FROM plniveducativo "
     Case "pfs"
      gdl_Procedure.ps_ReportName = "lstprofeoficio"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codpfs, despfs, estadopfs "
      s_Sql = s_Sql & "FROM plprofesion "
     Case "prd"
      gdl_Procedure.ps_ReportName = "lstperiodicidad"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codprd, desprd, estadoprd "
      s_Sql = s_Sql & "FROM plperiodicidad "
     Case "sec"
      gdl_Procedure.ps_ReportName = "lstseccion"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codsec, dessec, estadosec "
      s_Sql = s_Sql & "FROM plseccion "
     Case "sit"
      gdl_Procedure.ps_ReportName = "lstsitrapen"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codstp, desstp, estadostp "
      s_Sql = s_Sql & "FROM plsitrapen "
     Case "stp"
      gdl_Procedure.ps_ReportName = "lstsitrapen"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codstp, desstp, estadostp "
      s_Sql = s_Sql & "FROM plsitrapen "
     Case "tco"
      gdl_Procedure.ps_ReportName = "lsttipcontrato"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codtco, destco, estadotco "
      s_Sql = s_Sql & "FROM pltipcontrato "
     Case "tic"
      gdl_Procedure.ps_ReportName = "lsttipcom"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codtco, destco, estadotco "
      s_Sql = s_Sql & "FROM pltipcom "
     Case "tip"
      gdl_Procedure.ps_ReportName = "lsttippago"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codtip, destip, estadotip "
      s_Sql = s_Sql & "FROM pltippago "
     Case "tpt"
      gdl_Procedure.ps_ReportName = "lsttipotrabaja"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codtpt, destpt, estadotpt "
      s_Sql = s_Sql & "FROM pltpotrabajador "
     Case "tsu"
      gdl_Procedure.ps_ReportName = "lsttipsusp"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codtsu, destsu, estadotsu "
      s_Sql = s_Sql & "FROM pltipsusp "
     Case "ubica"
      gdl_Procedure.ps_ReportName = "lstubicacion"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codubica, desubica, estadoubica "
      s_Sql = s_Sql & "FROM plubicacion "
     Case "vfa"
      gdl_Procedure.ps_ReportName = "lstvinfami"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codvfa, desvfa, estadovfa "
      s_Sql = s_Sql & "FROM plvinfami "
     Case "via"
      gdl_Procedure.ps_ReportName = "lsttipovia"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codvia, desvia, abrevia, estadovia "
      s_Sql = s_Sql & "FROM pltipovia "
     Case "zona"
      gdl_Procedure.ps_ReportName = "lsttipozona"
      ' Genero sentencia de creacion de tabla temporal
      s_Sql = "SELECT codzona, deszona, abrezona, estadozona "
      s_Sql = s_Sql & "FROM pltipozona "
    End Select
   'ABRIL 2015
   ' s_Sql = s_Sql & "ORDER BY " & s_Orden & aElemento(0, 1)
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    Set porstRecordset = OpenRecordset(ps_StrgConnec & s_BaseDatos, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' ]
    
  End Select

End Sub
Private Sub dcaRegistro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  Select Case s_swEntidad
  Case "act"
    If FormVisible("fAbcActividad") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcActividad.ShowScreen
      End If
    End If
   Case "est"
    If FormVisible("fAbcEstablecimiento") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcEstablecimiento.ShowScreen
      End If
    End If
   Case "ldn"
    If FormVisible("fAbcLargaDistancia") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcLargaDistancia.ShowScreen
      End If
    End If
   Case "nac"
    If FormVisible("fAbcNacionalidad") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcNacionalidad.ShowScreen
      End If
    End If
   Case "niv"
    If FormVisible("fAbcNivEducativo") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcNivEducativo.ShowScreen
      End If
    End If
   Case "prd"
    If FormVisible("fAbcPeriodicidad") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcPeriodicidad.ShowScreen
      End If
    End If
   Case "stp"
    If FormVisible("fAbcSitrapen") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcSitrapen.ShowScreen
      End If
    End If
    
   Case "tco"
    If FormVisible("fAbcTipContrato") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipContrato.ShowScreen
      End If
    End If
   Case "bco"
    If FormVisible("fAbcEntidadBanco") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcEntidadBanco.ShowScreen
      End If
    End If
   Case "cdt"           ' condición de trabajo
    If FormVisible("fAbcCondicionTrabajo") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcCondicionTrabajo.ShowScreen
      End If
    End If
   Case "cgo"
    If FormVisible("fAbcCargoPersonal") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcCargoPersonal.ShowScreen
      End If
    End If
   Case "cls"
    If FormVisible("fAbcClasePlanilla") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcClasePlanilla.ShowScreen
      End If
    End If
   Case "cpc"
    If FormVisible("fAbcConcepCalculo") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcConcepCalculo.ShowScreen
      End If
    End If
   Case "dci"
    If FormVisible("fAbcDocIdentidad") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcDocIdentidad.ShowScreen
      End If
    End If
   Case "dmo"
    If FormVisible("fAbcDistribuMoneda") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcDistribuMoneda.ShowScreen
      End If
    End If
   Case "eps"
    If FormVisible("fAbcEntidadEps") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcEntidadEps.ShowScreen
      End If
    End If
   Case "pfs"
    If FormVisible("fAbcProfesion") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcProfesion.ShowScreen
      End If
    End If
   Case "sec"
    If FormVisible("fAbcSeccion") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcSeccion.ShowScreen
      End If
    End If
   Case "tpt"
    If FormVisible("fAbcTipoTrabajador") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipoTrabajador.ShowScreen
      End If
    End If
   Case "ubica"
    If FormVisible("fAbcUbicacion") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcUbicacion.ShowScreen
      End If
    End If
   Case "via"
    If FormVisible("fAbcTipoVia") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipoVia.ShowScreen
      End If
    End If
   Case "zona"
    If FormVisible("fAbcTipoZona") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipoZona.ShowScreen
      End If
    End If
   Case "sit"
    If FormVisible("fAbcSitrapen") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcSitrapen.ShowScreen
      End If
    End If
   Case "tip"
    If FormVisible("fAbcTipoPago") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipPago.ShowScreen
      End If
    End If
   Case "mof"
    If FormVisible("fAbcMotFin") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcMotFin.ShowScreen
      End If
    End If
   Case "mfo"
    If FormVisible("fAbcModForma") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcModForma.ShowScreen
      End If
    End If
  Case "vfa"
    If FormVisible("fAbcVinFami") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcVincFami.ShowScreen
      End If
    End If
Case "bdh"
    If FormVisible("fAbcMotBajaDH") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcMotBajaDH.ShowScreen
      End If
    End If
Case "tsu"
    If FormVisible("fAbcTipoSusp") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipSusp.ShowScreen
      End If
    End If
Case "tic"
    If FormVisible("fAbcTipoCom") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcTipComp.ShowScreen
      End If
    End If
Case "con"
    If FormVisible("fAbcSunat") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcSunat.ShowScreen
      End If
    End If
Case "cao"
    If FormVisible("fAbcCatOcu") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fabcCatOcu.ShowScreen
      End If
    End If
Case "cta"
    If FormVisible("fAbcCuenta") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcCuenta.ShowScreen
      End If
    End If
Case "ctr"
    If FormVisible("fAbcConven") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fabcConven.ShowScreen
      End If
    End If
End Select
End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6300: Me.Width = 6220
  Me.Left = 105: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = cmbTabla.Text
  
  ReDim aElemento(3, 10)
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 0) = Choose(i + 1, "Código", "Descripción", "Ok")
    aElemento(i, 1) = Choose(i + 1, "cod", "det", "estado")
    aElemento(i, 2) = Choose(i + 1, 800, 3556.03, 300)
    aElemento(i, 3) = Choose(i + 1, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(i, 4) = Choose(i + 1, "", "", "")
    aElemento(i, 5) = Choose(i + 1, False, False, False)
    aElemento(i, 6) = Choose(i + 1, True, True, True)
    aElemento(i, 7) = Choose(i + 1, "", "", "")
    aElemento(i, 8) = Choose(i + 1, dbgTop, dbgTop, dbgTop)
    aElemento(i, 9) = Choose(i + 1, 0, 0, 0)
  Next i
  ReDim aElementos(1, 3)
  For i = 0 To (UBound(aElementos, 1) - 1)
    aElementos(i, 0) = ""
    aElementos(i, 1) = 13427690: aElementos(i, 2) = vbBlack
  Next i
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(2).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(2).ValueItems.Translate = True
  For i = 0 To 1
    tdbRegistro.Columns(2).ValueItems.Add Item
    tdbRegistro.Columns(2).ValueItems.Item(i).Value = Choose(i + 1, s_Estado_Act, s_Estado_Ina)
    tdbRegistro.Columns(2).ValueItems.Item(i).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(i + 1, "estadok", "estadnok") & ".bmp")
  Next i
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(8, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "registro": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For i = 0 To (UBound(aElemento, 1) - 1)
      aElemento(i, 1) = Choose(i + 1, "seleccio", "anadir", "borrar", "ordascen", "orddesce", "busqueda", "prelimin", "Imprimir")
      aElemento(i, 2) = Choose(i + 1, "Selecciona y Edita " & s_TitleTable, "Añadir " & s_TitleTable, "Eliminar " & s_TitleTable, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Presentación Preliminar", "Imprimir")
      aElemento(i, 3) = Choose(i + 1, "&s", "&n", "&e", "&a", "&d", "&b", "&v", "&i")
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  
  s_swEntidad = "bco"   ' Inicializo el caso de la tabla
  ' Cargo los casos de tablas
  For i = 0 To 32
    cmbTabla.AddItem Choose(i + 1, "Tipo Actividad Empresarial", "Cargo de Personal", "Clase de Planilla", "Concepto de Cálculo", "Condición de Trabajo", "Cuenta Contable", "Distribución de Moneda", "Documento de Identidad", "Entidad Bancaria", "Entidad EPS", "Larga Distancia Nacional", "Modalidad Formativa", "Motivo de Baja de Derecho Habiente", "Motivo de Baja Trabajador", "Nacionalidad", "Situación Educativa", "Periodicidad Remuneración", "Profesión u Ocupación", "Sección de Empresa", "Situación Trabajador o Pensionista", "Tipo de Comprobante", "Tipo de Contrato", "Tipo Establecimientos Empresa", "Tipo de Pago", "Tipo Suspensión Relación Laboral", "Tipo de Vía", "Tipo de zona", "Tipo Trabajador", "Ubicación o Localidad", "Vinculo Familiar", "Conceptos Sunat", "Categoria Ocupacional", "Convenios x Tributación")
  Next i
  cmbTabla.ListIndex = 0
  
End Sub
Private Sub Form_Unload(Cancel As Integer)

  Select Case s_swEntidad
  Case "act"
    If FormVisible("fAbcActividad") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcActividad.Caption, vbExclamation
      Cancel = True
    End If
    Case "est"
    If FormVisible("fAbcEstablecimiento") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcEstablecimiento.Caption, vbExclamation
      Cancel = True
    End If
    Case "ldn"
    If FormVisible("fAbcLargaDistancia") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcLargaDistancia.Caption, vbExclamation
      Cancel = True
    End If
    Case "cta"
    If FormVisible("fAbcCuenta") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcCuenta.Caption, vbExclamation
      Cancel = True
    End If
    Case "nac"
    If FormVisible("fAbcNacionalidad") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcNacionalidad.Caption, vbExclamation
      Cancel = True
    End If
   Case "niv"
    If FormVisible("fAbcNivEducativo") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcNivEducativo.Caption, vbExclamation
      Cancel = True
    End If
   Case "tco"
    If FormVisible("fAbcTipContrato") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipContrato.Caption, vbExclamation
      Cancel = True
    End If
   Case "stp"
    If FormVisible("fAbcSitrapen") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcSitrapen.Caption, vbExclamation
      Cancel = True
    End If
   Case "prd"
    If FormVisible("fAbcPeriodicidad") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcPeriodicidad.Caption, vbExclamation
      Cancel = True
    End If
   Case "bco"
    If FormVisible("fAbcEntidadBanco") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcEntidadBanco.Caption, vbExclamation
      Cancel = True
    End If
   Case "cdt"           ' condición de trabajo
    If FormVisible("fAbcCondicionTrabajo") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcCondicionTrabajo.Caption, vbExclamation
      Cancel = True
    End If
   Case "cgo"
    If FormVisible("fAbcCargoPersonal") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcCargoPersonal.Caption, vbExclamation
      Cancel = True
    End If
   Case "cls"
    If FormVisible("fAbcClasePlanilla") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcClasePlanilla.Caption, vbExclamation
      Cancel = True
    End If
   Case "cpc"
    If FormVisible("fAbcConcepCalculo") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcConcepCalculo.Caption, vbExclamation
      Cancel = True
    End If
   Case "dci"
    If FormVisible("fAbcDocIdentidad") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcDocIdentidad.Caption, vbExclamation
      Cancel = True
    End If
   Case "dmo"
    If FormVisible("fAbcDistribuMoneda") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcDistribuMoneda.Caption, vbExclamation
      Cancel = True
    End If
   Case "eps"
    If FormVisible("fAbcEntidadEps") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcEntidadEps.Caption, vbExclamation
      Cancel = True
    End If
   Case "pfs"
    If FormVisible("fAbcProfesion") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcProfesion.Caption, vbExclamation
      Cancel = True
    End If
   Case "sec"
    If FormVisible("fAbcSeccion") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcSeccion.Caption, vbExclamation
      Cancel = True
    End If
   Case "tpt"
    If FormVisible("fAbcTipoTrabajador") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipoTrabajador.Caption, vbExclamation
      Cancel = True
    End If
   Case "ubica"
    If FormVisible("fAbcUbicacion") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcUbicacion.Caption, vbExclamation
      Cancel = True
    End If
   Case "via"
    If FormVisible("fAbcTipoVia") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipoVia.Caption, vbExclamation
      Cancel = True
    End If
   Case "zona"
    If FormVisible("fAbcTipoZona") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipoZona.Caption, vbExclamation
      Cancel = True
    End If
   Case "sit"
    If FormVisible("fAbcSitrapen") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcSitrapen.Caption, vbExclamation
      Cancel = True
    End If
    Case "tip"
    If FormVisible("fAbcTipPago") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipPago.Caption, vbExclamation
      Cancel = True
    End If
   Case "mof"
    If FormVisible("fAbcMotFin") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcMotFin.Caption, vbExclamation
      Cancel = True
    End If
   Case "mfo"
    If FormVisible("fAbcModForma") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcModForma.Caption, vbExclamation
      Cancel = True
    End If
   Case "vfa"
    If FormVisible("fAbcVincFami") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcVincFami.Caption, vbExclamation
      Cancel = True
    End If
   Case "bdh"
    If FormVisible("fAbcMotBajaDH") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcMotBajaDH.Caption, vbExclamation
      Cancel = True
    End If
   Case "tsu"
    If FormVisible("fAbcTipSusp") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipSusp.Caption, vbExclamation
      Cancel = True
    End If
   Case "tic"
    If FormVisible("fAbcTipComp") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcTipComp.Caption, vbExclamation
      Cancel = True
    End If
    Case "con"
    If FormVisible("fAbcSunat") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcSunat.Caption, vbExclamation
      Cancel = True
    End If
    Case "cao"
    If FormVisible("fAbcCatOcu") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fabcCatOcu.Caption, vbExclamation
      Cancel = True
    End If
    Case "ctr"
    If FormVisible("fAbcConven") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fabcConven.Caption, vbExclamation
      Cancel = True
    End If
  End Select
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
