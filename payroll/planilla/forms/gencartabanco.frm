VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fGeneraCartaBanco 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4020
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "gencartabanco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6420
   Begin TabDlg.SSTab tabRegister 
      Height          =   2850
      Left            =   75
      TabIndex        =   15
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5027
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
      TabPicture(0)   =   "gencartabanco.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblHelp(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "shpCuadro(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTexto(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTexto(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDato(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpFecha"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdHelp(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtRemunera"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNroCarta"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMotivo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtInteres"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtInteres 
         Height          =   300
         Left            =   4485
         TabIndex        =   7
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtMotivo 
         Height          =   300
         Left            =   1380
         TabIndex        =   9
         Top             =   1620
         Width           =   4650
      End
      Begin VB.TextBox txtNroCarta 
         Height          =   300
         Left            =   270
         TabIndex        =   3
         Top             =   1260
         Width           =   1590
      End
      Begin VB.TextBox txtRemunera 
         Height          =   300
         Left            =   1380
         TabIndex        =   11
         Top             =   1965
         Width           =   855
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   2310
         TabIndex        =   19
         Top             =   1965
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   2520
         TabIndex        =   5
         Top             =   1260
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   141819905
         CurrentDate     =   37515
      End
      Begin VB.Label lblDato 
         Caption         =   "Tasa de Interes :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   4485
         TabIndex        =   6
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
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
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   1
         Top             =   540
         Width           =   375
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
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
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   210
         Width           =   375
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   810
         Index           =   0
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   5985
      End
      Begin VB.Label lblDato 
         Caption         =   "Fecha de Proceso :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   945
         Width           =   1440
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   8
         Top             =   1665
         Width           =   1005
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
         Left            =   2670
         TabIndex        =   20
         Top             =   2010
         Width           =   195
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro de Carta :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   950
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Concepto :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   10
         Top             =   2010
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
         Left            =   5790
         TabIndex        =   17
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
         Picture         =   "gencartabanco.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   5400
         TabIndex        =   18
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "gencartabanco.frx":0044
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
         Left            =   285
         TabIndex        =   13
         Top             =   120
         Width           =   4800
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   14
      Top             =   3510
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   16
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
Attribute VB_Name = "fGeneraCartaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String, s_Mes As String           ' Codigo del registro, mes de transferencia
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Sub ShowScreen()
  Dim s_FechaCan As String
  
  s_FechaCan = gdl_Funcion.NumeroDiasMes(s_Mes, ps_Anyo) & "/" & s_Mes & "/" & ps_Anyo
  If fTransferBancos.ribAnalisis(1).Value Then
    s_Sql = "SELECT fechacan FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & "AND pdocts='" & Trim(fTransferBancos.txtPeriodo) & "'"
    s_Sql = s_Sql & "AND estadosub='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "ORDER BY subcts DESC"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    s_FechaCan = Format(porstRecordset!fechacan, s_FormatoFecha)
    s_Mes = Format(s_FechaCan, "mm")
  End If
    
  If Me.Tag = s_MdoData_Upd Then
    gdl_Procedure.EditText "AT", txtNroCarta, gdl_Funcion.aTexto(fTransferBancos.dcaRegistro.Recordset("nrocarta")), Me.Tag, False, fTransferBancos.dcaRegistro.Recordset("nrocarta").DefinedSize
    gdl_Procedure.EditDTPicker "AT", dtpFecha, fTransferBancos.dcaRegistro.Recordset("fechaproce"), Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txtInteres, FormatNumber(fTransferBancos.dcaRegistro.Recordset("porinteres"), 2), IIf(fTransferBancos.ribAnalisis(0).Value, s_MdoData_Vis, Me.Tag), False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo, gdl_Funcion.aTexto(fTransferBancos.dcaRegistro.Recordset("desmotivo")), Me.Tag, False, fTransferBancos.dcaRegistro.Recordset("desmotivo").DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera, gdl_Funcion.aTexto(fTransferBancos.dcaRegistro.Recordset("codcpc")), Me.Tag, False, fTransferBancos.dcaRegistro.Recordset("codcpc").DefinedSize
  Else
    gdl_Procedure.EditText "AT", txtNroCarta, "", Me.Tag, False, fTransferBancos.dcaRegistro.Recordset("nrocarta").DefinedSize
    gdl_Procedure.EditDTPicker "AT", dtpFecha, s_FechaCan, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txtInteres, FormatNumber(0, 2), IIf(fTransferBancos.ribAnalisis(0).Value, s_MdoData_Vis, Me.Tag), (fTransferBancos.ribAnalisis(0).Value), 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo, "", Me.Tag, False, fTransferBancos.dcaRegistro.Recordset("desmotivo").DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera, "", Me.Tag, False, fTransferBancos.dcaRegistro.Recordset("codcpc").DefinedSize
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera, "CP")
  lblTexto(0) = " Banco : " & Trim(fTransferBancos.txtBanco) & " - " & Trim(fTransferBancos.lblHelp(1)) & " "
  lblTexto(1) = " Periodo : " & Trim(fTransferBancos.txtPeriodo) & " - " & Trim(fTransferBancos.lblHelp(0)) & " "

End Sub
']
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Conceptos de planilla remuneracion
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Remuneración"
    s_Registro = ps_ClsPlanilla & "F" & "0"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + 800 + (cmdHelp(Index).Height / 2)
  tdbHelp.Left = tabRegister.Left + (1600 + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  Dim sSubPeriodo As String
  
  ' Realizo las validaciones de los campos a actualizar
  If txtNroCarta.Text = "" Then Beep: MsgBox "Debe Ingresar el numero de carta de Transfrencia", vbExclamation: txtNroCarta.SetFocus: Exit Sub
  If Not (Format(dtpFecha, "yyyymm") >= (ps_Anyo & s_Mes)) Then Beep: MsgBox "Fecha de proceso no valida; verificar", vbExclamation: dtpFecha.SetFocus: Exit Sub
  If txtMotivo.Text = "" Then Beep: MsgBox "Debe Ingresar el Motivo de la transfrencia", vbExclamation: txtMotivo.SetFocus: Exit Sub
  If txtRemunera.Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración", vbExclamation: txtRemunera.SetFocus: Exit Sub
  If lblHelp(0).Caption = "???" Then Beep: MsgBox "Concepto de Remuneración no es valido; Verificar", vbExclamation: txtRemunera.SetFocus: Exit Sub
  s_Registro = Trim(fTransferBancos.txtBanco.Text)
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
     
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Elimino los datos de carta existente
  s_Sql = "DELETE FROM plcartabanco "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo='" & Trim(fTransferBancos.txtPeriodo.Text) & "' "
  s_Sql = s_Sql & "AND codbco='" & s_Registro & "' "
  s_Sql = s_Sql & "AND nrocarta='" & Trim(txtNroCarta.Text) & "'"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
  
  If fTransferBancos.ribAnalisis(0).Value Then
    ' Inserto los registros
    s_Sql = "INSERT INTO plcartabanco ("
    s_Sql = s_Sql & "codcls, codbco, nrocarta, codcpc, desmotivo, codpsn, codpdo, fechaproce, "
    s_Sql = s_Sql & "codmon, importe_mn, importe_me, porinteres, usrcre, fyhcre, usrmdf, fyhmdf) "
    s_Sql = s_Sql & "SELECT res.codcls, psn.codbcopago, '" & Trim(txtNroCarta.Text) & "', "
    s_Sql = s_Sql & "res.codcpc, '" & Trim(txtMotivo.Text) & "', res.codpsn, res.codpdo, '" & Format(dtpFecha, s_FmtFechMysql_0) & "', "
    s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', 'E', 'N') AS codmon, res.importe_mn, res.importe_me, " & CDec(txtInteres.Text) & ", "
    s_Sql = s_Sql & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "', Null, NULL "
    s_Sql = s_Sql & "FROM  plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND psn.codbcopago='" & s_Registro & "' "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.codpdo='" & Trim(fTransferBancos.txtPeriodo) & "' "
    s_Sql = s_Sql & "AND res.codcpc='" & Trim(txtRemunera.Text) & "' "
    s_Sql = s_Sql & "ORDER BY codpsn"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Error
  Else
    ' Actualizo los datos de movimientos de C.T.S.
    s_Sql = "UPDATE plctsmovimiento mov, plpersonal psn "
    s_Sql = s_Sql & "SET mov.porinteres=0.00, mov.nrodeposito=Null "
    s_Sql = s_Sql & "WHERE mov.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND mov.pdocts='" & Trim(fTransferBancos.txtPeriodo.Text) & "' "
    s_Sql = s_Sql & "AND mov.nrodeposito='" & Trim(txtNroCarta.Text) & "' "
    s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "AND psn.codcls=mov.codcls "
    s_Sql = s_Sql & "AND psn.codpsn=mov.codpsn "
    s_Sql = s_Sql & "AND psn.codbcocts='" & s_Registro & "' "
    If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
    
    ' Selecciono el sub periodo mayor cancelado de liquidación de C.T.S.
    s_Sql = "SELECT DISTINCTROW MAX(CONCAT(res.pdoano,res.subcts)) AS subcts "
    s_Sql = s_Sql & "FROM plctsresultado res "
    s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdocts='" & Trim(fTransferBancos.txtPeriodo.Text) & "' "
    s_Sql = s_Sql & "AND res.codcpc='" & Trim(txtRemunera.Text) & "' "
    s_Sql = s_Sql & "AND psn.codbcocts='" & s_Registro & "' "
    s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "AND IFNULL(mov.nrodeposito, '')=''"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    sSubPeriodo = ""
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sSubPeriodo = Mid(gdl_Funcion.aTexto(porstRecordset!subcts), 5)
    End If
    porstRecordset.Close
    
    ' Inserto los registros
    s_Sql = "INSERT INTO plcartabanco ("
    s_Sql = s_Sql & "codcls, codbco, nrocarta, codcpc, desmotivo, codpsn, codpdo, fechaproce, "
    s_Sql = s_Sql & "codmon, importe_mn, importe_me, porinteres, usrcre, fyhcre, usrmdf, fyhmdf) "
    s_Sql = s_Sql & "SELECT res.codcls, psn.codbcocts, '" & Trim(txtNroCarta.Text) & "', "
    s_Sql = s_Sql & "res.codcpc, '" & Trim(txtMotivo.Text) & "', res.codpsn, res.pdocts, '" & Format(dtpFecha, s_FmtFechMysql_0) & "', "
    s_Sql = s_Sql & "IF(psn.ctsdolar='" & s_Estado_Act & "', 'E', 'N') AS codmon, res.importe_mn, res.importe_me, " & CDec(txtInteres.Text) & ", "
    s_Sql = s_Sql & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "', Null, NULL "
    s_Sql = s_Sql & "FROM plctsresultado res "
    s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdocts='" & Trim(fTransferBancos.txtPeriodo) & "' "
    s_Sql = s_Sql & "AND res.subcts='" & sSubPeriodo & "' "
    s_Sql = s_Sql & "AND res.codcpc='" & Trim(txtRemunera.Text) & "' "
    s_Sql = s_Sql & "AND psn.codbcocts='" & s_Registro & "' "
    s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "AND IFNULL(mov.nrodeposito, '')='' "
    s_Sql = s_Sql & "ORDER BY codpsn"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Error
    
    ' Actualizo los datos de movimientos de C.T.S.
    s_Sql = "UPDATE plctsmovimiento mov, plpersonal psn "
    s_Sql = s_Sql & "SET mov.porinteres=" & CDec(txtInteres.Text) & ", mov.nrodeposito='" & Trim(txtNroCarta.Text) & "' "
    s_Sql = s_Sql & "WHERE mov.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND mov.pdocts='" & Trim(fTransferBancos.txtPeriodo.Text) & "' "
    s_Sql = s_Sql & "AND mov.subcts='" & sSubPeriodo & "' "
    s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "AND IFNULL(mov.nrodeposito, '')='' "
    s_Sql = s_Sql & "AND psn.codcls=mov.codcls "
    s_Sql = s_Sql & "AND psn.codpsn=mov.codpsn "
    s_Sql = s_Sql & "AND psn.codbcocts='" & s_Registro & "' "
    If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se Actualizo exitosamente el " & lblTitle.Caption, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fTransferBancos.dcaRegistro, fTransferBancos.tdbRegistro, lblTitle.Caption
  ' Ubico el registro ingresado o actualizado
  fTransferBancos.dcaRegistro.Recordset.Find ("nrocarta='" & Trim(txtNroCarta.Text) & "'")
  If Not (fTransferBancos.dcaRegistro.Recordset.EOF Or fTransferBancos.dcaRegistro.Recordset.BOF) Then fTransferBancos.dcaRegistro.Recordset.MoveLast
  ShowScreen
  txtNroCarta.SetFocus
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 4500: Me.Width = 6510
  Me.Left = 3580: Me.Top = 2500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización de Carta de Transferencia de Bancos"
  lblTitle.Caption = "Parametros de Transferencia"
  s_Mes = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, fTransferBancos.txtPeriodo.Text, "MP")
  s_Mes = IIf(s_Mes = "???", ps_Mes, s_Mes)
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
    
  ' Obtengo el modo de operación del registro
  Me.Tag = fTransferBancos.Tag
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "edit": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "detactas"
  aElemento(0, 2) = "Actualizar Información de " & lblTitle.Caption
  gdl_Procedure.ViewGrafics Me, cmdUpdate, aElemento
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle.Caption
  cmdCancel.Cancel = True
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Cálculo", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Concepto de remuneración ganada
    txtRemunera = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtRemunera.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Conceptos de remuneración ganada
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 1     ' Concepto de quinta categoria
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Act
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 2     ' Tabla de UIT
    s_Sql = gdl_Funcion.HelpTablas("tbl", tdbHelp.Columns(ColIndex).DataField, ps_Anyo & ps_ClsPlanilla, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtInteres_GotFocus()
  gdl_Procedure.MarcaGet txtInteres
End Sub
Private Sub txtInteres_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtInteres_Validate(Cancel As Boolean)
  txtInteres.Text = IIf(Not IsNumeric(txtInteres.Text), 0, txtInteres.Text)
  If CDec(txtInteres.Text) < 0 Then MsgBox "Factor no puede ser negativo; Verifique", vbInformation: txtInteres.SetFocus: Exit Sub
  txtInteres.Text = FormatNumber(CDec(txtInteres.Text), 2)
End Sub
Private Sub txtMotivo_GotFocus()
  gdl_Procedure.MarcaGet txtMotivo
End Sub
Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNroCarta_GotFocus()
  gdl_Procedure.MarcaGet txtNroCarta
End Sub
Private Sub txtNroCarta_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_GotFocus()
  gdl_Procedure.MarcaGet txtRemunera
End Sub
Private Sub txtRemunera_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtRemunera_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera, "CP")
End Sub
