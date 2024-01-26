VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fConsultaProvision 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5775
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   11955
   Icon            =   "conprovision.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   11955
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11955
      _Version        =   65536
      _ExtentX        =   21087
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
         Left            =   10200
         TabIndex        =   4
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
         Picture         =   "conprovision.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   9840
         TabIndex        =   3
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "conprovision.frx":0028
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
         Left            =   870
         TabIndex        =   5
         Top             =   120
         Width           =   8085
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   5265
      Width           =   11955
      _Version        =   65536
      _ExtentX        =   21087
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
         Left            =   7200
         TabIndex        =   7
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
         Picture         =   "conprovision.frx":0044
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   6810
         TabIndex        =   8
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
         Picture         =   "conprovision.frx":0060
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   5100
         TabIndex        =   9
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
         Picture         =   "conprovision.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   4710
         TabIndex        =   10
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
         Picture         =   "conprovision.frx":0098
      End
      Begin MSAdodcLib.Adodc dcaRegistro 
         Height          =   330
         Left            =   90
         Top             =   90
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
   Begin Threed.SSFrame frmCuadro 
      Height          =   4680
      Left            =   30
      TabIndex        =   2
      Top             =   540
      Width           =   11050
      _Version        =   65536
      _ExtentX        =   19491
      _ExtentY        =   8255
      _StockProps     =   14
      ForeColor       =   0
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
      Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
         Height          =   4530
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   11000
         _ExtentX        =   19394
         _ExtentY        =   7990
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         BorderStyle     =   0
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
         CollapseColor   =   12632064
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&"
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
         _StyleDefs(34)  =   "Named:id=33:Normal"
         _StyleDefs(35)  =   ":id=33,.parent=0"
         _StyleDefs(36)  =   "Named:id=34:Heading"
         _StyleDefs(37)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(38)  =   ":id=34,.wraptext=-1"
         _StyleDefs(39)  =   "Named:id=35:Footing"
         _StyleDefs(40)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(41)  =   "Named:id=36:Selected"
         _StyleDefs(42)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(43)  =   "Named:id=37:Caption"
         _StyleDefs(44)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(45)  =   "Named:id=38:HighlightRow"
         _StyleDefs(46)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=39:EvenRow"
         _StyleDefs(48)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(49)  =   "Named:id=40:OddRow"
         _StyleDefs(50)  =   ":id=40,.parent=33"
         _StyleDefs(51)  =   "Named:id=41:RecordSelector"
         _StyleDefs(52)  =   ":id=41,.parent=34"
         _StyleDefs(53)  =   "Named:id=42:FilterBar"
         _StyleDefs(54)  =   ":id=42,.parent=33"
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4620
      Index           =   0
      Left            =   11160
      TabIndex        =   11
      Top             =   615
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8149
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
         Index           =   2
         Left            =   150
         TabIndex        =   13
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
         Picture         =   "conprovision.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   14
         Tag             =   "0"
         Top             =   1440
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
         Picture         =   "conprovision.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   15
         Tag             =   "0"
         Top             =   2040
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
         Picture         =   "conprovision.frx":00EC
      End
   End
End
Attribute VB_Name = "fConsultaProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer                              ' Indice para bucle
Private s_OptRegistro As String                         ' Instancia del formulario activo
Private o_Formulario As Object                          ' Objeto de la instancia del formulario activo
'[
Sub RecuperaRegistros()
    
  ' Recupera información
  If s_OptRegistro = "pvsvacacio" Then        ' Provisión de vacaciones
    s_Sql = "SELECT det.pdoano, det.pdomes, det.fechaini, det.fechafin, det.numerodias, det.diasvacper, det.diasvacacu, "
    s_Sql = s_Sql & "det.codmon, det.remunera_mn, det.remunera_me, det.importepvs_mn, det.importepvs_me, det.imporpvsacu_mn, "
    s_Sql = s_Sql & "det.imporpvsper_mn, det.imporpvsper_me, det.imporpvsacu_mn, det.imporpvsacu_me, "
    s_Sql = s_Sql & "det.fechacan, det.estadodet, det.codcta_debmn, det.codcta_habmn, det.codcta_debme, det.codcta_habme "
    s_Sql = s_Sql & "FROM plpvsvacaciondet det "
    s_Sql = s_Sql & "WHERE det.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & "AND det.codpvs='" & o_Formulario.dcaRegistro.Recordset!codpvs & "' "
    s_Sql = s_Sql & "AND det.codpsn='" & o_Formulario.dcaRegistro.Recordset!codpsn & "' "
    s_Sql = s_Sql & "AND det.pdopvs='" & o_Formulario.dcaRegistro.Recordset!pdopvs & "' "
    s_Sql = s_Sql & "ORDER BY det.pdoano, det.pdomes"
  ElseIf s_OptRegistro = "pvsgratifi" Then        ' Provisión de gratificaciones
    s_Sql = "SELECT gra.pdoano, gra.pdomes, gra.fechaini, gra.fechafin, gra.numerodias, "
    s_Sql = s_Sql & "gra.codmon, gra.remunera_mn, gra.remunera_me, gra.imporpvsacu_mn, "
    s_Sql = s_Sql & "gra.imporpvsacu_me, gra.importepvs_mn, gra.importepvs_me, gra.fechacan, gra.estadogra "
    s_Sql = s_Sql & "FROM plpvsgratifica gra "
    s_Sql = s_Sql & "WHERE gra.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & "AND gra.pdoano='" & o_Formulario.dcaRegistro.Recordset!pdoano & "' "
    s_Sql = s_Sql & "AND gra.sempvs='" & o_Formulario.dcaRegistro.Recordset!sempvs & "' "
    s_Sql = s_Sql & "AND gra.codpsn='" & o_Formulario.dcaRegistro.Recordset!codpsn & "'"
    s_Sql = s_Sql & "ORDER BY gra.pdomes"
  End If
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbRegistro.DataSource = porstRecordset

End Sub
']
Private Sub cmdAction_Click(Index As Integer)
  
  Select Case Index
   Case 0, 1
    ' Verifico que Existan Registros
    If tdbRegistro.VisibleRows = 0 Then Beep: MsgBox "No Existen Registros para Imprimir", vbExclamation: Exit Sub
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = UCase$(lblTitle)
    gdl_Procedure.ps_ReportName = IIf(s_OptRegistro = "pvsgratifi", "cstpvsgratifi", "cstpvsvacacion")
    ReDim aElemento(2, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    If s_OptRegistro = "pvsvacacio" Then        ' Provisión de vacaciones
      aElemento(2, 2) = "Periodo;" & "Provisión - " & Trim(o_Formulario.dcaRegistro.Recordset!codpvs) & ";true"
    ElseIf s_OptRegistro = "pvsgratifi" Then        ' Provisión de gratificaciones
      aElemento(2, 2) = "Periodo;" & Trim(o_Formulario.dcaRegistro.Recordset!pdoano) & " - Semestre " & Trim(o_Formulario.dcaRegistro.Recordset!sempvs) & ";true"
    End If
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    If s_OptRegistro = "pvsvacacio" Then        ' Provisión de vacaciones
      s_Sql = "SELECT vac.codpsn, vac.pdoano, vac.pdomes, vac.pdopvs, pdo.descripvs, sub.fechaini AS subfechaini, sub.fechafin AS subfechafin, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, vac.fechaini, vac.fechafin, vac.fechacan, "
      s_Sql = s_Sql & "vac.numerodias, vac. diasvacper, vac.diasvacacu, vac.codmon, vac.remunera_mn, vac.remunera_me, "
      s_Sql = s_Sql & "vac.importepvs_mn, vac.importepvs_me, vac.imporpvsper_mn, vac.imporpvsper_me, vac.imporpvsacu_mn, vac.imporpvsacu_me "
      s_Sql = s_Sql & "FROM plpvsvacaciondet vac "
      s_Sql = s_Sql & "INNER JOIN plpvsperiodovac pdo ON vac.codcls=pdo.codcls AND vac.codpvs=pdo.codpvs "
      s_Sql = s_Sql & "INNER JOIN plpvsvacacion sub ON vac.codcls=sub.codcls AND vac.codpvs=sub.codpvs AND vac.codpsn=sub.codpsn AND vac.pdopvs=sub.pdopvs "
      s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON vac.codcls=psn.codcls AND vac.codpsn=psn.codpsn "
      s_Sql = s_Sql & "WHERE vac.codcls='" & ps_ClsPlanilla & "'"
      s_Sql = s_Sql & "AND vac.codpvs='" & o_Formulario.dcaRegistro.Recordset!codpvs & "' "
      s_Sql = s_Sql & "AND vac.codpsn='" & o_Formulario.dcaRegistro.Recordset!codpsn & "' "
      s_Sql = s_Sql & "AND vac.pdopvs='" & o_Formulario.dcaRegistro.Recordset!pdopvs & "' "
      s_Sql = s_Sql & "ORDER BY vac.pdoano, " & aElemento(0, 1)
    ElseIf s_OptRegistro = "pvsgratifi" Then        ' Provisión de gratificaciones
      s_Sql = "SELECT gra.codpsn, gra.pdoano, gra.sempvs, gra.pdomes, pdo.descripvs, pdo.mesini, pdo.mesfin, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, psn.codcco, cco.detcco, "
      s_Sql = s_Sql & "gra.fechaini, gra.fechafin, gra.numerodias, gra.codmon, gra.remunera_mn, gra.remunera_me, gra.imporpvsacu_mn, "
      s_Sql = s_Sql & "gra.imporpvsacu_me , gra.importepvs_mn, gra.importepvs_me, gra.fechacan "
      s_Sql = s_Sql & "FROM plpvsgratifica gra "
      s_Sql = s_Sql & "INNER JOIN plpvsperiodogra pdo ON gra.codcls=pdo.codcls AND gra.pdoano=pdo.pdoano AND gra.sempvs=pdo.sempvs "
      s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON gra.codcls=psn.codcls AND gra.codpsn=psn.codpsn "
      s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON cco.codcco=psn.codcco "
      s_Sql = s_Sql & "WHERE gra.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND gra.pdoano='" & o_Formulario.dcaRegistro.Recordset!pdoano & "' "
      s_Sql = s_Sql & "AND gra.sempvs='" & o_Formulario.dcaRegistro.Recordset!sempvs & "' "
      s_Sql = s_Sql & "AND gra.codpsn='" & o_Formulario.dcaRegistro.Recordset!codpsn & "'"
      s_Sql = s_Sql & "ORDER BY gra.sempvs, " & aElemento(0, 1)
    End If
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, Index, False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
   Case 2
    quecodpvs = fPvsVacacion.tdbRegistro.Columns(0).Text
    quecodpsn = o_Formulario.dcaRegistro.Recordset!codpsn
    quepdopvs = Replace(fPvsVacacion.tdbRegistro.Columns(2).Text, "-", "")
    quepdoano = tdbRegistro.Columns(0).Text
    queopcion = "A"
    
    fAbcVacacionAnt.Show
   Case 3
    If tdbRegistro.Columns(0).Text = "" Then
      Exit Sub
    End If
    
    quecodpvs = fPvsVacacion.tdbRegistro.Columns(0).Text
    quecodpsn = o_Formulario.dcaRegistro.Recordset!codpsn
    quepdopvs = Replace(fPvsVacacion.tdbRegistro.Columns(2).Text, "-", "")
    quepdoano = tdbRegistro.Columns(0).Text
    
    quepdomes = tdbRegistro.Columns(1).Text
    quefechaini = tdbRegistro.Columns(2).Text
    quefechafin = tdbRegistro.Columns(3).Text
    quenumerodias = tdbRegistro.Columns(4).Text
    quecodmon = tdbRegistro.Columns(17).Text
    queremunera_mn = tdbRegistro.Columns(5).Text
    queremunera_me = tdbRegistro.Columns(6).Text
    queimporpvsacu_mn = tdbRegistro.Columns(7).Text
    queimporpvsacu_me = tdbRegistro.Columns(8).Text
    queimportepvs_mn = tdbRegistro.Columns(9).Text
    queimportepvs_me = tdbRegistro.Columns(10).Text
    quefechacan = tdbRegistro.Columns(12).Text
    queestadodet = Right(tdbRegistro.Columns(11).Text, 1)
    quecodcta_debmn = tdbRegistro.Columns(13).Text
    quecodcta_habmn = tdbRegistro.Columns(14).Text
    quecodcta_debme = tdbRegistro.Columns(15).Text
    quecodcta_habme = tdbRegistro.Columns(16).Text
    
    queopcion = "C"
    fAbcVacacionAnt.Show
    fAbcVacacionAnt.dtpFechas(0).Enabled = True
    fAbcVacacionAnt.dtpFechas(1).Enabled = True
   Case 4
    If tdbRegistro.Columns(0).Text = "" Then
      Exit Sub
    End If
    
    If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
      ' Coloco el puntero en espera
      gdl_Procedure.PunteroEnEspera
      '[ Inicio la conexión a la base de datos ]
      ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
      ' Creo los arreglos de eliminacion
      a_Where = Array("codcls", "codpvs", "codpsn", "pdopvs", "pdoano", "pdomes")
    
    Dim quemes As String
    
      Select Case tdbRegistro.Columns(1).Text
       Case "Enero"
        quemes = "01"
       Case "Febrero"
        quemes = "02"
       Case "Marzo"
        quemes = "03"
       Case "Abril"
        quemes = "04"
       Case "Mayo"
        quemes = "05"
       Case "Junio"
        quemes = "06"
       Case "Julio"
        quemes = "07"
       Case "Agosto"
        quemes = "08"
       Case "Setiembre"
        quemes = "09"
       Case "Octubre"
        quemes = "10"
       Case "Noviembre"
        quemes = "11"
       Case "Diciembre"
        quemes = "12"
      End Select
    
      a_Valores = Array(ps_ClsPlanilla, fPvsVacacion.tdbRegistro.Columns(0).Text, o_Formulario.dcaRegistro.Recordset!codpsn, Replace(fPvsVacacion.tdbRegistro.Columns(2).Text, "-", ""), tdbRegistro.Columns(0).Text, quemes)
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      gdl_Conexion.IniciaTransaccion    'Inicia transacción
      ' Elimino el registro
      If Not Records_Del("plpvsvacaciondet", a_Where, a_Valores, a_Tipos) Then GoTo Error
      gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
      MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
      ' Refresco el ado control y la grilla
      ' fConsultaCalculo.RecuperaRegistros
      Unload Me
    End If
    GoTo Finalizar
    
Error:
    gdl_Conexion.CancelaTransaccion
Finalizar:
    ' Coloco el puntero en normal
    gdl_Procedure.PunteroNormal
    '[ Finalizo la conexión a la base de datos ]
    Set gdl_Conexion = Nothing
    If Not l_ExistRecord Then Unload Me
  End Select
  
End Sub
Private Sub cmdMove_Click(Index As Integer)
  
  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: o_Formulario.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not o_Formulario.dcaRegistro.Recordset.BOF Then o_Formulario.dcaRegistro.Recordset.MovePrevious
           If o_Formulario.dcaRegistro.Recordset.BOF Then o_Formulario.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not o_Formulario.dcaRegistro.Recordset.EOF Then o_Formulario.dcaRegistro.Recordset.MoveNext
           If o_Formulario.dcaRegistro.Recordset.EOF Then o_Formulario.dcaRegistro.Recordset.MoveLast
   Case 3: o_Formulario.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub Form_Load()
  Dim Item As New ValueItem   ' Cambio el formato de la grilla columna de valores
  
  'Establece posición y titulo del formulario
  Me.Height = 6150
  Me.Left = 2500: Me.Top = 550
  
  ' Instancio el objeto
  s_OptRegistro = fMenu.Tag
  If s_OptRegistro = "pvsvacacio" Then
    Set o_Formulario = fPvsVacacion
    lblTitle = "Provisión de Vacaciones"
    Me.Width = 12030
  ElseIf s_OptRegistro = "pvsgratifi" Then
    Set o_Formulario = fPvsGratificacion
    lblTitle = "Provisión de Gratificaciones"
    Me.Width = 11235
  End If
  ' Titulo del formulario y panel
  s_TitleWindow = "Resultados de Proceso de Cálculo"
      
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(5, 2)
  ' Icono y título del formulario
  aElemento(5, 1) = "reporte": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 4
    aElemento(n_Index, 1) = Choose(n_Index + 1, "prelimin", "Imprimir", "seleccio", "anadir", "borrar")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Presentación Preliminar", "Imprimir", "Selecciona y Edita " & lblTitle, "Añadir " & lblTitle, "Eliminar " & lblTitle)
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "reporte": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  '[ Configuración de la grilla de consulta
  If s_OptRegistro = "pvsvacacio" Then        ' Provisión de vacaciones
    ReDim aElemento(22, 10)
    For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Periodo", "Mes", "Fec.Inicio", "Fec.Final", "Días", "Día.Per", "Día.Acu", "Remun " & s_Codmon_mn_Txt, "Pvs Mes " & s_Codmon_mn_Txt, "Pvs.Peri." & s_Codmon_mn_Txt, "Pvs.Acum." & s_Codmon_mn_Txt, "Remun " & s_Codmon_me_Txt, "Pvs Mes " & s_Codmon_me_Txt, "Pvs.Peri." & s_Codmon_me_Txt, "Pvs.Acum." & s_Codmon_me_Txt, "Ok", "Cancelacion", "cuentadebmn", "cuentahabmn", "cuentadebme", "cuentahabme", "moneda")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "pdoano", "pdomes", "fechaini", "fechafin", "numerodias", "diasvacper", "diasvacacu", "remunera_mn", "importepvs_mn", "imporpvsper_mn", "imporpvsacu_mn", "remunera_me", "importepvs_me", "imporpvsper_me", "imporpvsacu_me", "estadodet", "fechacan", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme", "codmon")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 450, 820, 920, 920, 600, 600, 600, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 300, 0, 0, 0, 0, 0, 0)
      aElemento(n_Index, 3) = Choose(n_Index + 1, vbCenter, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbCenter, vbCenter, vbCenter, vbCenter, vbCenter, vbCenter, vbCenter)
      aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", s_FormatoFecha, s_FormatoFecha, "standard", "standard", "standard", "standard", "standard", "standard", "standard", "standard", "standard", "standard", "standard", "", s_FormatoFecha, "", "", "", "", "standard")
      'aElemento(n_Index, 5) = Choose(n_Index + 1, dbgMergeRestricted, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
      aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
      aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)
      aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
      aElemento(n_Index, 8) = Choose(n_Index + 1, dbgCenter, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
      aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 0, 3, 3, 3, 3, 3, 3)
    Next n_Index
  ElseIf s_OptRegistro = "pvsgratifi" Then        ' Provisión de gratificaciones
    ReDim aElemento(17, 10)
    For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Mes", "Fec.Inicio", "Fec.Final", "Dias", "Remun " & s_Codmon_mn_Txt, "Pvs.Acum." & s_Codmon_mn_Txt, "Pvs Mes " & s_Codmon_mn_Txt, "Remun " & s_Codmon_me_Txt, "Pvs.Acum." & s_Codmon_me_Txt, "Pvs Mes " & s_Codmon_me_Txt, "Ok", "Extra " & s_Codmon_mn_Txt, "Pvs.Acum." & s_Codmon_mn_Txt, "Pvs Mes " & s_Codmon_mn_Txt, "Extra " & s_Codmon_me_Txt, "Pvs.Acum." & s_Codmon_me_Txt, "Pvs Mes " & s_Codmon_me_Txt)
      aElemento(n_Index, 1) = Choose(n_Index + 1, "pdomes", "fechaini", "fechafin", "numerodias", "remunera_mn", "imporpvsacu_mn", "importepvs_mn", "remunera_me", "imporpvsacu_me", "importepvs_me", "estadogra", "remunera_mn", "imporpvsacu_mn", "importepvs_mn", "remunera_me", "imporpvsacu_me", "importepvs_me")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 980, 950, 950, 500, 1200, 1200, 1200, 1200, 1200, 1200, 300, 1200, 1200, 1200, 1200, 1200, 1200)
      aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbCenter, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify)
      aElemento(n_Index, 4) = Choose(n_Index + 1, "", s_FormatoFecha, s_FormatoFecha, "", "standard", "standard", "standard", "standard", "standard", "standard", "", "standard", "standard", "standard", "standard", "standard", "standard")
      aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
      aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True)
      aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
      aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
      aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 2, 2, 2, 2, 2, 2)
    Next n_Index
  End If
  ReDim aElementos(3, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = n_BackColorMdf: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, "", 3
  ' Cambio el formato de la grilla columna de valores
  If s_OptRegistro = "pvsvacacio" Then        ' Provisión de vacaciones
    ' Cambio el formato de la grilla columna de valores
    tdbRegistro.Columns(1).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(1).ValueItems.Translate = True
    For n_Index = 0 To 11
      ' Estado del registro
      tdbRegistro.Columns(1).ValueItems.Add Item
      tdbRegistro.Columns(1).ValueItems.Item(n_Index).Value = Format(n_Index + 1, "00")
      tdbRegistro.Columns(1).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    Next n_Index
    tdbRegistro.Columns(15).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(15).ValueItems.Translate = True
    For n_Index = 0 To 2
      ' Estado del registro
      tdbRegistro.Columns(15).ValueItems.Add Item
      tdbRegistro.Columns(15).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Ina, s_Estado_Act, s_Estado_Blq)
      tdbRegistro.Columns(15).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "periogen", "periopvs", "periocan") & ".bmp")
    Next n_Index
  ElseIf s_OptRegistro = "pvsgratifi" Then        ' Provisión de gratificaciones
    ' Cambio el formato de la grilla columna de valores
    tdbRegistro.Columns(0).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(0).ValueItems.Translate = True
    For n_Index = 0 To 11
      ' Estado del registro
      tdbRegistro.Columns(0).ValueItems.Add Item
      tdbRegistro.Columns(0).ValueItems.Item(n_Index).Value = Format(n_Index + 1, "00")
      tdbRegistro.Columns(0).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    Next n_Index
    
    tdbRegistro.Columns(10).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(10).ValueItems.Translate = True
    For n_Index = 0 To 2
      ' Estado del registro
      tdbRegistro.Columns(10).ValueItems.Add Item
      tdbRegistro.Columns(10).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Ina, s_Estado_Act, s_Estado_Blq)
      tdbRegistro.Columns(10).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "periogen", "periopvs", "periocan") & ".bmp")
    Next n_Index
  End If
  ']
  ' Carga los datos en el formulario
  RecuperaRegistros
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  
  If s_OptRegistro = "consulxcpc" Then
    ' Restauro los valores del periodo seleccionado
    o_Formulario.txtPeriodo.Locked = False
    o_Formulario.txtPeriodo.BackColor = &HFFFFFF
    o_Formulario.txtPeriodo.ForeColor = &HC00000
    o_Formulario.cmdHelp(0).Enabled = True
  End If
  Set o_Formulario = Nothing
  
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then tdbRegistro.DataSource = Nothing: RecuperaRegistros
End Sub
