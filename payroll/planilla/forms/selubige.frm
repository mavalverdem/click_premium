VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSeleccionUbigeo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   3675
   ClientTop       =   3465
   ClientWidth     =   6180
   Icon            =   "selubige.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6180
      _Version        =   65536
      _ExtentX        =   10901
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
         Index           =   0
         Left            =   4880
         TabIndex        =   15
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
         Picture         =   "selubige.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   5270
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
         Picture         =   "selubige.frx":0028
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
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   4200
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6150
      _Version        =   65536
      _ExtentX        =   10848
      _ExtentY        =   4313
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
      Begin VB.TextBox txtUbigeo 
         Height          =   300
         Index           =   2
         Left            =   165
         MaxLength       =   7
         TabIndex        =   10
         Top             =   1935
         Width           =   960
      End
      Begin VB.TextBox txtUbigeo 
         Height          =   300
         Index           =   1
         Left            =   210
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1245
         Width           =   960
      End
      Begin VB.TextBox txtUbigeo 
         Height          =   300
         Index           =   0
         Left            =   210
         MaxLength       =   3
         TabIndex        =   2
         Top             =   570
         Width           =   960
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   3
         Top             =   570
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   1
         Left            =   1230
         TabIndex        =   7
         Top             =   1245
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   2
         Left            =   1185
         TabIndex        =   11
         Top             =   1935
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   1530
         TabIndex        =   12
         Top             =   1965
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Distrito"
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
         Index           =   2
         Left            =   225
         TabIndex        =   9
         Top             =   1635
         Width           =   1365
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   8
         Top             =   1290
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Provincia"
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
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   960
         Width           =   1365
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   1575
         TabIndex        =   4
         Top             =   615
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Departamento"
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
         Left            =   225
         TabIndex        =   1
         Top             =   285
         Width           =   1365
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2175
      Left            =   1290
      TabIndex        =   16
      Top             =   1515
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   3836
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1535"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1931"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=37,.parent=2,.namedParent=39"
      _StyleDefs(17)  =   "FilterBarStyle:id=40,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=38,.parent=37"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=41,.parent=40"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(38)  =   "Named:id=29:Normal"
      _StyleDefs(39)  =   ":id=29,.parent=0"
      _StyleDefs(40)  =   "Named:id=30:Heading"
      _StyleDefs(41)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=30,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=31:Footing"
      _StyleDefs(44)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=32:Selected"
      _StyleDefs(46)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=33:Caption"
      _StyleDefs(48)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(49)  =   "Named:id=34:HighlightRow"
      _StyleDefs(50)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(51)  =   "Named:id=35:EvenRow"
      _StyleDefs(52)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=36:OddRow"
      _StyleDefs(54)  =   ":id=36,.parent=29"
      _StyleDefs(55)  =   "Named:id=39:RecordSelector"
      _StyleDefs(56)  =   ":id=39,.parent=30"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=29"
   End
   Begin MSAdodcLib.Adodc dcaHelp 
      Height          =   330
      Left            =   0
      Top             =   3405
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
      Caption         =   "Adodc1"
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
Attribute VB_Name = "fSeleccionUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo del Formulario
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private i As Integer, s_Filtro  As String               ' Indice para bucle
Private Sub InicializaCriterio(ByVal n_Index As Integer)
  If n_Index = (txtUbigeo.Count - 1) Then Exit Sub
  For i = (n_Index + 1) To (txtUbigeo.Count - 1)
    lblHelp(i) = "": txtUbigeo(i) = ""
  Next i

End Sub
Private Sub cmdAction_Click(Index As Integer)

  If Index = 0 Then
    If txtUbigeo(2) = "" Or txtUbigeo(2) = "???" Then txtUbigeo(2) = ""
    o_SwSelUbica.txtUbigeo(n_SwSelUbica) = txtUbigeo(2)
    o_SwSelUbica.lblUbigeo(n_SwSelUbica) = lblHelp(0) & "/" & lblHelp(1) & "/" & lblHelp(2)
  End If
  ' Cierro el formulario
  Unload Me

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  For i = 0 To Index - 1
    If txtUbigeo(i) = "" Or lblHelp(i) = "" Or lblHelp(i) = "???" Then Beep: MsgBox "Debe Ingresar o no existe el " & Choose(i + 1, "Pais", "Departamento", "Provincia"), vbCritical: txtUbigeo(i).SetFocus: Exit Sub
  Next i
  If n_IndexHelp = Index And Index <> 1 Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If
  
  s_Filtro = ""
  If Index > 0 Then
    s_Filtro = txtUbigeo(Index - 1)
  End If
  ' Recupero la información
  tdbHelp.Caption = "Ubicación Geografica - " & Choose(Index + 1, "Departamento", "Provincia", "Distrito")
  s_Sql = gdl_Funcion.HelpTablas("ubg", tdbHelp.Columns(0).DataField, Index, s_Filtro)
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_BDSystems, dcaHelp, tdbHelp, s_Sql, adCmdText, adLockReadOnly
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = frmCuadro.Top + (cmdHelp(0).Top - (cmdHelp(0).Height / 2)) + 200
  tdbHelp.Left = frmCuadro.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2450: tdbHelp.Width = 4600
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 4040: Me.Width = 6270
  Me.Left = 3930: Me.Top = 4150
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Selección de Ubicación Geografica"
  lblTitle = "Ubicación Geografica"
  n_IndexHelp = -1: s_Filtro = ""
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "seleccio": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For i = 0 To (UBound(aElemento, 1) - 1)
      aElemento(i, 1) = Choose(i + 1, "aceptar", "cancelar")
      aElemento(i, 2) = Choose(i + 1, "Acepta Información ", "Cancela Información ") & "Selección"
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(1).Cancel = True
  
  For i = 0 To 2: cmdHelp(i).Enabled = True: Next i
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For i = 0 To (UBound(aElemento, 1) - 1)
    aElemento(i, 0) = Choose(i + 1, "Código", "Descripción")
    aElemento(i, 1) = Choose(i + 1, "codubg", "desubg")
    aElemento(i, 2) = Choose(i + 1, 934.7402, 3350.764)
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
      aElementos(i, 1) = n_BackColorHelp#: aElementos(i, 2) = vbBlack
  Next i
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbHelp, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Ubicacion Geografica", 2
  ' Asigno el control de datos  ala grilla
  tdbHelp.DataSource = dcaHelp
  
  ' Recupero la información
  s_Sql = gdl_Funcion.HelpTablas("ubg", tdbHelp.Columns(0).DataField, 0, "")
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_BDSystems, dcaHelp, tdbHelp, s_Sql, adCmdText, adLockReadOnly
  ']

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal
   
End Sub
Private Sub tdbHelp_DblClick()

  If (dcaHelp.Recordset.EOF And dcaHelp.Recordset.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  ' ubicacion geografica
  txtUbigeo(n_IndexHelp) = tdbHelp.Columns(0).Value
  lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
  txtUbigeo(n_IndexHelp).SetFocus

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  s_Sql = gdl_Funcion.HelpTablas("ubg", tdbHelp.Columns(ColIndex).DataField, n_IndexHelp & s_Filtro, "")
  dcaHelp.RecordSource = s_Sql
  dcaHelp.Refresh

End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaHelp, tdbHelp, ""
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim porstClone As ADODB.Recordset, v_Bookmark As Variant
  Dim n_Columna As Integer, s_Criterio As String

  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    ' Conformo la cadena de ayuda
    s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    
    Set porstClone = dcaHelp.Recordset.Clone()
    n_Columna = tdbHelp.Col
    s_Criterio = tdbHelp.Columns(n_Columna).DataField & " >= '" & s_SqlHelp & "'"
    porstClone.Find s_Criterio, 0, adSearchForward, 0
    If Not (porstClone.BOF Or porstClone.EOF) Then
      dcaHelp.Recordset.Bookmark = porstClone.Bookmark
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
Private Sub txtUbigeo_Change(Index As Integer)
  InicializaCriterio Index
End Sub
Private Sub txtUbigeo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtUbigeo(Index)
End Sub
Private Sub txtUbigeo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtUbigeo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtUbigeo_LostFocus(Index As Integer)
  lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, Index, txtUbigeo(Index), "UB")
End Sub
