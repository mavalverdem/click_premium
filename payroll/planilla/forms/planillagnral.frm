VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fPlanillaGnral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "planillagnral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5835
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4600
      Left            =   45
      TabIndex        =   11
      Top             =   120
      Width           =   4950
      _ExtentX        =   8731
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
      Top             =   4770
      Width           =   4950
      _ExtentX        =   8731
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
      Left            =   5040
      TabIndex        =   0
      Top             =   120
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
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Tag             =   "0"
         Top             =   945
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
         Picture         =   "planillagnral.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1365
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
         Picture         =   "planillagnral.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Tag             =   "0"
         Top             =   2055
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
         Picture         =   "planillagnral.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   2475
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
         Picture         =   "planillagnral.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   3555
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
         Picture         =   "planillagnral.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   3975
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
         Picture         =   "planillagnral.frx":0098
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   10
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
         Top             =   510
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
         Picture         =   "planillagnral.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   2895
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
         Picture         =   "planillagnral.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   4395
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
         Picture         =   "planillagnral.frx":00EC
      End
   End
End
Attribute VB_Name = "fPlanillaGnral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, i As Integer            ' Indice de la barra de herramientas, indice para bucle
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Informaci�n
  s_Sql = "SELECT DISTINCTROW codpll, despll, sizefont, sizepapel, posipapel, imprimecab "
  s_Sql = s_Sql & " FROM plplanilla"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly

End Sub
Private Sub cmdAction_Click(Index As Integer)

  ' Inicializo el modo de registro o selecci�n
  Me.Tag = ""
  Select Case Index
   Case 0, 2  ' Visualizar o analizar, eliminar registro
    If Not (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Then
      Me.Tag = IIf(Index = 0, s_MdoData_Vis, s_MdoData_Del)
      fAbcPlanillaGnral.Show
    End If
   Case 1 ' Nuevo registro
    Me.Tag = s_MdoData_Ins
    fAbcPlanillaGnral.Show
   Case 3, 4  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index - 2, " ASC", " DESC")
   Case 5 ' Busqueda de registro
    If Not (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Then
      Set go_tdbBusqueda = tdbRegistro
      Set go_dcaBusqueda = dcaRegistro
      gn_ColBusqueda = tdbRegistro.Columns.Count
      fBusqueda.Show vbModal
    End If
   Case 6, 7 ' Opciones de impresi�n
    ' Verifico que Existan Registros
    If dcaRegistro.Recordset.RecordCount = 0 Then Beep: MsgBox "No Existen " & s_TitleTable & " para Imprimir", vbExclamation: Exit Sub
    ' Parametros de Impresi�n
    gdl_Procedure.ps_ReportTitle = "FORMATO DE PLANILLAS DE REMUNERACIONES"
    gdl_Procedure.ps_ReportName = "lstformaplla"
    ReDim aElemento(3, 2): ReDim aElementos(2)
    ' Parametros del store procedure
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generaci�n e impresi�n de informaci�n para el reporte
    s_Sql = "SELECT pll.codpll , pll.despll, pll.imprimecab, pll.fila, pll.columna, pll.tipo, pll.alias, "
    s_Sql = s_Sql & "pll.descripcion, pll.posicion, pll.longitud, pll.subrayado, "
    s_Sql = s_Sql & "IF(pll.tipo='C', cpc.descpc, var.descripcion) AS detalle "
    s_Sql = s_Sql & "FROM plplanilla pll "
    s_Sql = s_Sql & "LEFT JOIN plvarfunc var ON pll.tipo=var.tipo AND pll.alias=var.codigo AND var.tipo='D' "
    s_Sql = s_Sql & "LEFT JOIN plconcepto cpc ON pll.alias=cpc.codcpc AND pll.tipo='C' "
    s_Sql = s_Sql & "WHERE pll.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "ORDER BY codpll, fila, columna"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la informaci�n
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' ]
   Case 8  ' Impresi�n de cabeceras de planilla
    ' Verifico que Existan Registros
    If dcaRegistro.Recordset.RecordCount = 0 Then Beep: MsgBox "No Existen " & s_TitleTable & " para Imprimir", vbExclamation: Exit Sub
    fPrnTituloPlanilla.Show vbModal
  End Select

End Sub
Private Sub dcaRegistro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  If FormVisible("fAbcPlanillaGnral") Then
    If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
      fAbcPlanillaGnral.ShowScreen
    End If
  End If

End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

  ' Establece posici�n del formulario
  Me.Height = 5610: Me.Width = 5920
  Me.Left = 2750: Me.Top = 710
  ' Recupera par�metro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Registro de Planilla"
  s_TitleTable = "Formato de Planilla"
  
  ReDim aElemento(2, 10)
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 0) = Choose(i + 1, "C�digo", "Descripci�n")
    aElemento(i, 1) = Choose(i + 1, "codpll", "despll")
    aElemento(i, 2) = Choose(i + 1, 800, 3556.03)
    aElemento(i, 3) = Choose(i + 1, vbLeftJustify, vbLeftJustify)
    aElemento(i, 4) = Choose(i + 1, "", "")
    aElemento(i, 5) = Choose(i + 1, False, False)
    aElemento(i, 6) = Choose(i + 1, True, True)
    aElemento(i, 7) = Choose(i + 1, "", "")
    aElemento(i, 8) = Choose(i + 1, dbgTop, dbgTop)
    aElemento(i, 9) = Choose(i + 1, 0, 0)
  Next i
  ReDim aElementos(1, 3)
  For i = 0 To (UBound(aElementos, 1) - 1)
    aElementos(i, 0) = ""
    aElementos(i, 1) = 13427690: aElementos(i, 2) = vbBlack
  Next i
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupaci�n"
  
  ' Configuro parametros de visualizaci�n del formulario y los controles
  ReDim aElemento(9, 2)
  ' Icono y t�tulo del formulario
  aElemento(UBound(aElemento, 1), 1) = "registro": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 1) = Choose(i + 1, "seleccio", "anadir", "borrar", "ordascen", "orddesce", "busqueda", "prelimin", "Imprimir", "imprapid")
    aElemento(i, 2) = Choose(i + 1, "Selecciona y Edita " & s_TitleTable, "A�adir " & s_TitleTable, "Eliminar " & s_TitleTable, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Presentaci�n Preliminar", "Imprimir", "Imprime Cabecera")
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  
End Sub
Private Sub Form_Unload(Cancel As Integer)

  If FormVisible("fAbcPlanillaGnral") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fAbcPlanillaGnral.Caption, vbExclamation
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
