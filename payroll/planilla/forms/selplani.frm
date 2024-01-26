VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPlanilla 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "selplani.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
      Height          =   3720
      Left            =   90
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6562
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3254"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
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
      DeadAreaBackColor=   14737632
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
   Begin Threed.SSPanel panToolBar 
      Height          =   4095
      Index           =   0
      Left            =   5970
      TabIndex        =   0
      Top             =   120
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   7223
      _StockProps     =   15
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
         TabIndex        =   6
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
         Left            =   165
         TabIndex        =   1
         Tag             =   "0"
         Top             =   450
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
         Picture         =   "selplani.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   160
         TabIndex        =   2
         Tag             =   "0"
         Top             =   1160
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
         Picture         =   "selplani.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   160
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1870
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
         Picture         =   "selplani.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   160
         TabIndex        =   4
         Tag             =   "0"
         Top             =   2580
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
         Picture         =   "selplani.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Cancel          =   -1  'True
         Height          =   360
         Index           =   4
         Left            =   160
         TabIndex        =   5
         Tag             =   "0"
         Top             =   3290
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
         Picture         =   "selplani.frx":007C
      End
   End
   Begin MSAdodcLib.Adodc dcaSeleccion 
      Height          =   330
      Left            =   90
      Top             =   3900
      Width           =   5820
      _ExtentX        =   10266
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
      ForeColor       =   64
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
      Caption         =   ""
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
Attribute VB_Name = "fSelPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventana y grilla
Private n_IndexTool As Integer, i As Byte               ' Indice de la barra de herramientas, indice para bucle
'[
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, descls, clave, horadiaria "
  s_Sql = s_Sql & "FROM plclasplan "
  s_Sql = s_Sql & "WHERE estadocls='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaSeleccion, tdbSeleccion, s_Sql, adCmdText, adLockReadOnly

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_Clave As String, n_Index As Integer
    
  Select Case Index
   Case 0 ' Selecciona registro
    If gdl_Funcion.aTexto(dcaSeleccion.Recordset!clave) <> "" Then
      ' clave encriptada
      Inputbox_Password fSelPlanilla
      s_Clave = InputBox("Ingrese Clave de Acceso para la Clase " & Trim(dcaSeleccion.Recordset!descls), "Clave de Acceso")
      If s_Clave = "" Then Exit Sub
      If gdl_Funcion.Desencripta(dcaSeleccion.Recordset!clave) <> s_Clave Then
        Beep
        MsgBox "Clave de Acceso para la Clase " & Trim(dcaSeleccion.Recordset!descls) & " No es Correcta", vbCritical
        Exit Sub
      End If
    End If
    pl_Salir = True
  
    ' Capturo los valores de la planilla
    ps_ClsPlanilla = gdl_Funcion.aTexto(dcaSeleccion.Recordset!codcls)
    ps_DesClsPlanilla = UCase(gdl_Funcion.aTexto(dcaSeleccion.Recordset!descls))
    pn_HoroLaboraxDia = CDec(dcaSeleccion.Recordset!horadiaria)
    ' Nivel de centro de costo
    s_Sql = "SELECT nivelcencosto FROM plcfgempresa WHERE pdoano='" & ps_Anyo & "'"
    
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    pn_NivelCenCosto = 5
    If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      pn_NivelCenCosto = CInt(porstRecordset!nivelcencosto)
    End If
    Set porstRecordset = Nothing
  
    Unload Me
   Case 1, 2  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbSeleccion.Columns(tdbSeleccion.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 3 ' Busqueda de registro
    If Not (dcaSeleccion.Recordset.EOF Or dcaSeleccion.Recordset.BOF) Then
      Set go_tdbBusqueda = tdbSeleccion
      Set go_dcaBusqueda = dcaSeleccion
      gn_ColBusqueda = tdbSeleccion.Columns.Count
      fBusqueda.Show vbModal
    End If
   Case 4 ' Salir del Sistema
    Unload Me
  End Select

End Sub
Private Sub dcaSeleccion_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  ' Controla Movimiento de Puntero a BOF o EOF
  If Not (dcaSeleccion.Recordset.BOF And dcaSeleccion.Recordset.EOF) Then
    If dcaSeleccion.Recordset.BOF Then dcaSeleccion.Recordset.MoveFirst
    If dcaSeleccion.Recordset.EOF Then dcaSeleccion.Recordset.MoveLast
  End If

End Sub
Private Sub Form_Load()

  ' Establece posición del formulario
  gdl_Procedure.CentraFormulario Me
  gdl_Procedure.pl_RecordSelector = False: pl_Salir = False
    
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Selección de Clase de Planilla"
  s_TitleTable = "Planilla"
  
  ReDim aElemento(2, 10)
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 0) = Choose(i + 1, "Código", "Clase de Planilla")
    aElemento(i, 1) = Choose(i + 1, "codcls", "descls")
    aElemento(i, 2) = Choose(i + 1, 905.2599, 4555.168)
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
    aElementos(i, 0) = "": aElementos(i, 1) = 13427690
    aElementos(i, 2) = vbBlack
  Next i
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbSeleccion, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbSeleccion, s_TitleTable, 1
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(5, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "seleccio": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 1) = Choose(i + 1, "seleccio", "ordascen", "orddesce", "busqueda", "escapar")
    aElemento(i, 2) = Choose(i + 1, "Selecciona Registro", "Ordenar Ascendente", "Ordenar Descendente", "Buscar Registro", "Salir")
    aElemento(i, 3) = Choose(i + 1, "&s", "&a", "&d", "&b", "&s")
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  'Recupero los registros del control de datos asignado
  tdbSeleccion.DataSource = dcaSeleccion
  RecuperaRegistros tdbSeleccion.Columns(0).DataField & " ASC"

End Sub
Private Sub Form_Unload(Cancel As Integer)
  gdl_Procedure.pl_RecordSelector = True
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub tdbSeleccion_DblClick()
  ' Boton de selección de Registro
  cmdAction_Click 0
End Sub
Private Sub tdbSeleccion_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Refresca la información
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaSeleccion, tdbSeleccion, " " & s_TitleTable
End Sub
Private Sub tdbSeleccion_KeyPress(KeyAscii As Integer)
  ' Tecla enter boton de selección de Registro
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
