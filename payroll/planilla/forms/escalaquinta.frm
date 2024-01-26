VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fEscalaQuinta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "escalaquinta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   7560
   Begin Threed.SSFrame frmCuadro 
      Height          =   2850
      Left            =   15
      TabIndex        =   6
      Top             =   600
      Width           =   6780
      _Version        =   65536
      _ExtentX        =   11959
      _ExtentY        =   5027
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
         Height          =   2700
         Left            =   45
         TabIndex        =   7
         Top             =   120
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   4763
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         BorderStyle     =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
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
   Begin Threed.SSPanel panToolBar 
      Height          =   2775
      Index           =   0
      Left            =   6825
      TabIndex        =   1
      Top             =   675
      Width           =   690
      _Version        =   65536
      _ExtentX        =   1217
      _ExtentY        =   4895
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
         TabIndex        =   4
         Top             =   15
         Width           =   660
         _Version        =   65536
         _ExtentX        =   1164
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
         Left            =   120
         TabIndex        =   2
         Tag             =   "0"
         Top             =   1425
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
         Picture         =   "escalaquinta.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1845
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
         Picture         =   "escalaquinta.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "0"
         Top             =   615
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
         Picture         =   "escalaquinta.frx":0044
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   585
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7560
      _Version        =   65536
      _ExtentX        =   13335
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "P"
      ForeColor       =   12582912
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Font3D          =   1
   End
End
Attribute VB_Name = "fEscalaQuinta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private a_EscalaQuinta As New XArrayDB                  ' Array de remuneraciones por default
'[
Private Sub RecuperaRegistros(ByVal s_Orden As String)
  Dim nLimEscala(4) As Double
'MODIFICACION 7-ENERO-2015
'  s_Sql = "SELECT orden, numerouit, factor, IFNULL(tbl.valordefa, 0.00) AS valoruit "
'  s_Sql = s_Sql & "FROM plescalaquinta "
'  s_Sql = s_Sql & "LEFT JOIN plcfgempresa cfg ON cfg.pdoano='" & ps_Anyo & "' "
'  s_Sql = s_Sql & "LEFT JOIN pltablabase tbl ON tbl.codcls='" & ps_ClsPlanilla & "' AND cfg.pdoano=tbl.pdoano AND cfg.codtbluit=tbl.codtbl "
'  s_Sql = s_Sql & "ORDER BY " & s_Orden

  s_Sql = "SELECT orden, numerouit, factor, IFNULL(tbl.valordefa, 0.00) AS valoruit "
  s_Sql = s_Sql & "FROM plescalaquinta peq, plcfgempresa cfg, pltablabase tbl "
  s_Sql = s_Sql & "WHERE peq.pdoanyo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND cfg.pdoano=tbl.pdoano  AND cfg.codtbluit=tbl.codtbl "
  s_Sql = s_Sql & "AND cfg.pdoano=peq.pdoanyo AND tbl.codcls='" & ps_ClsPlanilla & "'  "
  s_Sql = s_Sql & "ORDER BY " & s_Orden
 

  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  a_EscalaQuinta.ReDim 1, 0, 0, 4
  ' Si hay registros  de remuneraciones
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    porstRecordset.MoveLast: porstRecordset.MoveFirst
    a_EscalaQuinta.ReDim 1, porstRecordset.RecordCount, 0, 4
    For n_Index = 1 To 4: nLimEscala(n_Index) = 0: Next n_Index
    n_Index = 0
    While Not porstRecordset.EOF
      n_Index = n_Index + 1
      nLimEscala(1) = nLimEscala(2)
      nLimEscala(2) = (CInt(porstRecordset!numerouit) * CDec(porstRecordset!valoruit))
      nLimEscala(3) = (nLimEscala(2) - nLimEscala(1)) * CDec(porstRecordset!factor)
      nLimEscala(4) = nLimEscala(4) + nLimEscala(3)
      a_EscalaQuinta(n_Index, 0) = Round(nLimEscala(1), 2)
      a_EscalaQuinta(n_Index, 1) = Round(nLimEscala(2), 2)
      a_EscalaQuinta(n_Index, 2) = CDec(porstRecordset!factor * 100)
      a_EscalaQuinta(n_Index, 3) = Round(nLimEscala(3), 2)
      a_EscalaQuinta(n_Index, 4) = Round(nLimEscala(4), 2)
      porstRecordset.MoveNext
    Wend
  End If
  ' Cierro el recordset y saco del entorno
  porstRecordset.Close: Set porstRecordset = Nothing
  
  ' Asigno el arreglo a la grilla y relleno la misma
  Set tdbRegistro.Array = a_EscalaQuinta
  tdbRegistro.Rebind

End Sub
Private Sub cmdAction_Click(Index As Integer)
  
  Select Case Index
   Case 0 ' Registro de parametro de uit
    RecuperaRegistros "orden ASC"
   Case 1, 2  ' Opciones de impresión
    ' Verifico que Existan Registros
    If (tdbRegistro.EOF And tdbRegistro.BOF) Or (tdbRegistro.VisibleRows = 0) Then Beep: MsgBox "No Existen " & tdbRegistro.Caption, vbExclamation: Exit Sub
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = Me.Caption
    gdl_Procedure.ps_ReportName = "cstescala5ta"
    
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "":  aElemento(1, 2) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & UCase(panToolBar(1).Caption) & ";true"
    aElemento(2, 2) = "Periodo;" & ps_Anyo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "orden smallint(2) Not Null, "
    s_Sql = s_Sql & "limiteini decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "limitefin decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "tasa decimal(6,2) Null Default '0', "
    s_Sql = s_Sql & "parcial decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "acumulado decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "PRIMARY KEY (orden)) "
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Inserto los valores
    For n_Index = 1 To a_EscalaQuinta.UpperBound(1)
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " "
      s_Sql = s_Sql & "VALUES("
      s_Sql = s_Sql & CInt(n_Index) & ", "
      s_Sql = s_Sql & CDec(a_EscalaQuinta(n_Index, 0)) & ", "
      s_Sql = s_Sql & CDec(a_EscalaQuinta(n_Index, 1)) & ", "
      s_Sql = s_Sql & CDec(a_EscalaQuinta(n_Index, 2)) & ", "
      s_Sql = s_Sql & CDec(a_EscalaQuinta(n_Index, 3)) & ", "
      s_Sql = s_Sql & CDec(a_EscalaQuinta(n_Index, 4)) & ")"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    Next n_Index
        
    ' Selecciono la información del reporte
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM  tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY orden"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 1), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' ]
  End Select

End Sub
Private Sub Form_Load()

  ' Establece posición del formulario
  Me.Height = 3960: Me.Width = 7650
  Me.Top = 800: Me.Left = 2500
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Desde", "Hasta", "T %", "Impto. Parcial", "Impto. Acumulado")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "inicio", "final", "factor", "parcial", "acumulado")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1262, 1262, 500, 1400, 1650)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify, vbRightJustify)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "standard", "standard", "standard", "standard", "standard")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0)
  Next n_Index
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, "Escala de Impuesto", 3
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(3, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "eleccion": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "promedio", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Parametro de Escala", "Presentación Preliminar", "Imprimir")
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&p", "&v", "&i")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  panToolBar(1).Caption = "          Escala de Impuestos para Retenciones y Pagos a Cuenta          Del Impuesto a la Renta de la Quinta Categoria"
 
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  RecuperaRegistros "orden ASC"
  
End Sub

Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index
  
End Sub

