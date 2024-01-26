VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fConsultaCalculo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5775
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   8550
   Icon            =   "concalculo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8550
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      _Version        =   65536
      _ExtentX        =   15081
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
         Left            =   7665
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
         Picture         =   "concalculo.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   7275
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
         Picture         =   "concalculo.frx":0028
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   195
         TabIndex        =   5
         Top             =   120
         Width           =   6645
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   5265
      Width           =   8550
      _Version        =   65536
      _ExtentX        =   15081
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
         Left            =   5400
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
         Picture         =   "concalculo.frx":0044
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   5010
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
         Picture         =   "concalculo.frx":0060
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   3300
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
         Picture         =   "concalculo.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2910
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
         Picture         =   "concalculo.frx":0098
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   4685
      Left            =   30
      TabIndex        =   2
      Top             =   540
      Width           =   8490
      _Version        =   65536
      _ExtentX        =   14975
      _ExtentY        =   8264
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
         Height          =   4535
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   8415
         _ExtentX        =   14843
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
End
Attribute VB_Name = "fConsultaCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private s_OptRegistro As String                         ' Instancia del formulario activo
Private o_Formulario As Object                          ' Objeto de la instancia del formulario activo
'[
Sub RecuperaRegistros()
  
  ' Recupera información
  If s_OptRegistro = "consulxcpc" Then            ' Planilla general
    lblTitle = o_Formulario.dcaRegistro.Recordset!codpsn & " - " & o_Formulario.dcaRegistro.Recordset!apepaterno & " " & o_Formulario.dcaRegistro.Recordset!apematerno & " " & o_Formulario.dcaRegistro.Recordset!nombres
    s_Sql = "SELECT res.codcpc, cpc.descpc, res.tipocpc, cxp.defaultcpc, "
    s_Sql = s_Sql & "cxp.clasecpc, res.impbolecpc, res.importe_mn, res.importe_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN plconceplanilla cxp ON res.codcls=cxp.codcls AND res.codcpc=cxp.codcpc "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codpdo='" & Trim(o_Formulario.txtPeriodo) & "' "
    s_Sql = s_Sql & "AND res.codpsn='" & o_Formulario.dcaRegistro.Recordset!codpsn & "' "
    s_Sql = s_Sql & "ORDER BY secuencia"
  ElseIf s_OptRegistro = "consulxpsn" Then        ' Conceptos por Persona
    lblTitle = o_Formulario.dcaRegistro.Recordset!codcpc & " - " & o_Formulario.dcaRegistro.Recordset!descpc
    s_Sql = "SELECT res.codpsn, cpc.descpc, psn.codpsn, "
    s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
    s_Sql = s_Sql & "(CASE codmon "
    s_Sql = s_Sql & "WHEN 'N' THEN '" & s_Codmon_mn_Txt & "' "
    s_Sql = s_Sql & "WHEN 'E' THEN '" & s_Codmon_me_Txt & "' END) AS codmon, res.importe_mn, res.importe_me "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codpdo='" & o_Formulario.txtPeriodo & "' "
    s_Sql = s_Sql & "AND res.codcpc='" & o_Formulario.dcaRegistro.Recordset!codcpc & "' "
    s_Sql = s_Sql & "ORDER BY res.codcpc, res.codpsn"
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
    If s_OptRegistro = "consulxcpc" Then            ' concepto por
      gdl_Procedure.ps_ReportTitle = "Resultado de Cálculo de Concepto por Personal"
      gdl_Procedure.ps_ReportName = "cstconcextraba"
    ElseIf s_OptRegistro = "consulxpsn" Then        ' Copceptos por persona
      gdl_Procedure.ps_ReportTitle = "Resultado de Cálculo de Personal por Concepto"
      gdl_Procedure.ps_ReportName = "cstconceptosxpsn"
    End If
  
    ReDim aElemento(2, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & UCase(gdl_Procedure.ps_ReportTitle) & ";true"
    aElemento(2, 2) = "Periodo;" & Trim(o_Formulario.txtPeriodo.Text) & " - " & Trim(o_Formulario.lblHelp(0).Caption) & ";true"
    
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    If s_OptRegistro = "consulxcpc" Then            ' Planilla general
      s_Sql = "SELECT res.codpsn, CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
      s_Sql = s_Sql & "res.secuencia, res.codcpc, cpc.descpc , res.tipocpc, cxp.defaultcpc, cxp.clasecpc, res.impbolecpc, res.importe_mn, res.importe_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plconceplanilla cxp ON res.codcls=cxp.codcls AND res.codcpc=cxp.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & Trim(o_Formulario.txtPeriodo) & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & o_Formulario.dcaRegistro.Recordset!codpsn & "' "
      s_Sql = s_Sql & "ORDER BY codpsn, secuencia, " & aElemento(0, 1)
    ElseIf s_OptRegistro = "consulxpsn" Then        ' Copceptos por Persona
      s_Sql = "SELECT res.codcpc, cpc.descpc, psn.codpsn, CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
      s_Sql = s_Sql & "(CASE codmon "
      s_Sql = s_Sql & "WHEN 'N' THEN '" & s_Codmon_mn_Txt & "' "
      s_Sql = s_Sql & "WHEN 'E' THEN '" & s_Codmon_me_Txt & "' END) AS moneda, res.importe_mn, res.importe_me "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & o_Formulario.txtPeriodo & "' "
      s_Sql = s_Sql & "AND res.codcpc='" & o_Formulario.dcaRegistro.Recordset!codcpc & "' "
      s_Sql = s_Sql & "ORDER BY res.codcpc, res.codpsn"
    End If
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, Index, False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
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
  Me.Height = 6150: Me.Width = 8640
  Me.Left = 2500: Me.Top = 550
  
  ' Instancio el objeto
  s_OptRegistro = fMenu.Tag
  If s_OptRegistro = "consulxcpc" Then
    Set o_Formulario = o_SelConsulxcpc
    lblTitle = o_Formulario.tdbRegistro.Columns(0).Text & " - " & o_Formulario.tdbRegistro.Columns(1).Text & " " & o_Formulario.tdbRegistro.Columns(2).Text & " " & o_Formulario.tdbRegistro.Columns(3).Text
  ElseIf s_OptRegistro = "consulxpsn" Then
    Set o_Formulario = o_SelConsulxpsn
    lblTitle = o_Formulario.tdbRegistro.Columns(1).Text & " - " & o_Formulario.tdbRegistro.Columns(2).Text
  End If
  ' Bloqueo los valores del periodo seleccionado
  o_Formulario.txtPeriodo.Locked = True
  o_Formulario.txtPeriodo.BackColor = &HC7D8E0
  o_Formulario.txtPeriodo.ForeColor = &HC00000
  o_Formulario.cmdHelp(0).Enabled = False
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Resultados de Proceso de Cálculo"
      
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(2, 1) = "reporte": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 1
    aElemento(n_Index, 1) = Choose(n_Index + 1, "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Presentación Preliminar", "Imprimir")
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
  If s_OptRegistro = "consulxcpc" Then            ' Planilla general
    ReDim aElemento(8, 10)
    For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Concepto", "Descripción", "Clase", "Tipo", "Default", "Importe " & s_Codmon_mn_Txt, "Importe " & s_Codmon_me_Txt, "Prn")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcpc", "descpc", "clasecpc", "tipocpc", "defaultcpc", "importe_mn", "importe_me", "impbolecpc")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 880, 2395, 500, 730, 500, 1200, 1200, 500)
      aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbCenter, vbCenter, vbCenter, vbRightJustify, vbRightJustify, vbCenter)
      aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "", "standard", "standard", "")
      aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False, False, False, False)
      aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True, True, True)
      aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "", "", "", "")
      aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
      aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0, 0, 0, 0)
    Next n_Index
    ReDim aElementos(1, 3)
    For n_Index = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Index, 0) = ""
      aElementos(n_Index, 1) = n_BackColorMdf: aElementos(n_Index, 2) = vbBlack
    Next n_Index
  ElseIf s_OptRegistro = "consulxpsn" Then            ' Conceptos x Persona
    ReDim aElemento(5, 10)
    For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Codigo", "Nombre", "Moneda", "Importe " & s_Codmon_mn_Txt, "Importe " & s_Codmon_me_Txt)
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "apellidosnombres", "codmon", "importe_mn", "importe_me")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 950, 3750, 800, 1200, 1200)
      aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbCenter, vbRightJustify, vbRightJustify)
      aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "standard", "standard")
      aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False)
      aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True)
      aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "")
      aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
      aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0)
    Next n_Index
    ReDim aElementos(1, 3)
    For n_Index = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Index, 0) = ""
      aElementos(n_Index, 1) = n_BackColorMdf: aElementos(n_Index, 2) = vbBlack
    Next n_Index
  End If
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, "", 3
  ' Cambio el formato de la grilla columna de valores
  If s_OptRegistro = "consulxcpc" Then            ' Planilla general
    tdbRegistro.Columns(3).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(3).ValueItems.Translate = True
    For n_Index = 0 To 2
      tdbRegistro.Columns(3).ValueItems.Add Item
      tdbRegistro.Columns(3).ValueItems.Item(n_Index).Value = Trim(n_Index)
      tdbRegistro.Columns(3).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, "Ingreso", "Descuento", "Aporte")
    Next n_Index
    
    tdbRegistro.Columns(2).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(2).ValueItems.Translate = True
    tdbRegistro.Columns(4).ValueItems.Presentation = dbgCheckBox
    tdbRegistro.Columns(7).ValueItems.Presentation = dbgCheckBox
    For n_Index = 0 To 1
      ' Calse de concepto
      tdbRegistro.Columns(2).ValueItems.Add Item
      tdbRegistro.Columns(2).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "C", "F")
      tdbRegistro.Columns(2).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "constant", "funcion") & ".bmp")
    Next n_Index
  ElseIf s_OptRegistro = "consulxpsn" Then
    ' Cambio el formato de la grilla columna de valores para la moneda
    tdbRegistro.Columns(2).ValueItems.Presentation = dbgNormal
    tdbRegistro.Columns(2).ValueItems.Translate = True
    For n_Index = 0 To 1
      tdbRegistro.Columns(2).ValueItems.Add Item
      tdbRegistro.Columns(2).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Codmon_mn, s_Codmon_me)
      tdbRegistro.Columns(2).ValueItems.Item(n_Index).DisplayValue = Choose(n_Index + 1, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    Next n_Index
  End If
  ']
  ' Carga los datos en el formulario
  RecuperaRegistros
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  
  If (s_OptRegistro = "consulxcpc" Or s_OptRegistro = "consulxpsn") Then
    ' Restauro los valores del periodo seleccionado
    o_Formulario.txtPeriodo.Locked = False
    o_Formulario.txtPeriodo.BackColor = &HFFFFFF
    o_Formulario.txtPeriodo.ForeColor = &HC00000
    o_Formulario.cmdHelp(0).Enabled = True
    o_Formulario.lblDato(0).Tag = ""
  End If
  Set o_Formulario = Nothing
  
End Sub
