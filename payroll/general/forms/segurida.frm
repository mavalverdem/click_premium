VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form fSeguridad 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6255
   ClientLeft      =   2265
   ClientTop       =   540
   ClientWidth     =   7275
   Icon            =   "segurida.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   7275
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7275
      _Version        =   65536
      _ExtentX        =   12832
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
         Left            =   5025
         TabIndex        =   11
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "segurida.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   5415
         TabIndex        =   12
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "segurida.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   6225
         TabIndex        =   13
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "segurida.frx":0044
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
         Left            =   195
         TabIndex        =   10
         Top             =   120
         Width           =   4305
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   4875
      Index           =   1
      Left            =   45
      TabIndex        =   4
      Top             =   1335
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
      _ExtentY        =   8599
      _StockProps     =   14
      Caption         =   " Opciones Asignadas "
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
      Font3D          =   3
      ShadowStyle     =   1
      Begin Threed.SSPanel panContenedor 
         Height          =   4530
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   285
         Width           =   3450
         _Version        =   65536
         _ExtentX        =   6085
         _ExtentY        =   7990
         _StockProps     =   15
         Caption         =   "Empresa"
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
         BevelOuter      =   1
         Alignment       =   6
         Begin MSOutl.Outline outEmpresa 
            Height          =   4500
            Left            =   15
            TabIndex        =   6
            Top             =   15
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   7937
            _StockProps     =   77
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
            MouseIcon       =   "segurida.frx":0060
            PicturePlus     =   "segurida.frx":007C
            PictureMinus    =   "segurida.frx":0176
            PictureLeaf     =   "segurida.frx":0270
            PictureOpen     =   "segurida.frx":0792
            PictureClosed   =   "segurida.frx":088C
         End
      End
      Begin Threed.SSPanel panContenedor 
         Height          =   4530
         Index           =   1
         Left            =   3615
         TabIndex        =   7
         Top             =   300
         Width           =   3450
         _Version        =   65536
         _ExtentX        =   6085
         _ExtentY        =   7990
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
         BevelOuter      =   1
         Begin MSOutl.Outline outSeguridad 
            Height          =   4500
            Left            =   15
            TabIndex        =   8
            Top             =   15
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   7937
            _StockProps     =   77
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
            MouseIcon       =   "segurida.frx":0DAE
            PicturePlus     =   "segurida.frx":0DCA
            PictureMinus    =   "segurida.frx":0EC4
            PictureLeaf     =   "segurida.frx":0FBE
            PictureOpen     =   "segurida.frx":14E0
            PictureClosed   =   "segurida.frx":15DA
         End
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   705
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
      _ExtentY        =   1244
      _StockProps     =   14
      Caption         =   " Usuario "
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      ShadowStyle     =   1
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         MaxLength       =   10
         TabIndex        =   1
         Top             =   285
         Width           =   1260
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   1575
         TabIndex        =   2
         Top             =   285
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
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
         Left            =   1935
         TabIndex        =   3
         Top             =   330
         Width           =   195
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Bindings        =   "segurida.frx":1AFC
      Height          =   2175
      Left            =   1560
      TabIndex        =   14
      Top             =   495
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
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
End
Attribute VB_Name = "fSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle y parametro de codigo
Private s_OldUsuario As String                          ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private a_Empresa() As String, a_Opciones() As String   ' Arreglos de seguridad
'[
Private Sub Empresa_Opcion()
  Dim nContador As Integer
  
  ' Obtengo las empresas
  s_Sql = "SELECT codemp, razemp FROM tgemp ORDER BY codemp, razemp"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Reconfiguro registro de empresas
  ReDim a_Empresa(1)
  outEmpresa.AddItem "Registro de Empresas"
  outEmpresa.Indent(0) = 0
  outEmpresa.PictureType(0) = outOpen
  a_Empresa(outEmpresa.ListCount - 1) = "empresa"
  ' Cargo las empresas
  While Not porstRecordset.EOF
    outEmpresa.ListIndex = -1
    outEmpresa.AddItem porstRecordset("codemp") & " - " & porstRecordset("razemp")
    ' Redimensiono el array de empresas
    ReDim Preserve a_Empresa(UBound(a_Empresa, 1) + 1)
    a_Empresa(outEmpresa.ListCount - 1) = porstRecordset("codemp")
    porstRecordset.MoveNext
    DoEvents
  Wend
  porstRecordset.Close
  
  ' Reconfiguro las opciones del sistema
  ReDim a_Opciones(1)
  outSeguridad.AddItem " " & ps_NomSistema
  outSeguridad.Indent(0) = 0
  outSeguridad.PictureType(0) = outOpen
  a_Opciones(outSeguridad.ListCount) = "opcion"
  ' Obtengo las opciones del sistema
  s_Sql = "SELECT codmdl, opcion, orden, detmdl, RIGHT(IF(opcion='0', orden, opcion), 1) AS item, "
  s_Sql = s_Sql & "IF(opcion='0', opcion, orden) AS secuencia "
  s_Sql = s_Sql & "FROM sgmdl "
  s_Sql = s_Sql & "WHERE codsis='" & ps_CodSistema & "' "
  s_Sql = s_Sql & "AND detmdl<>'-' "
  s_Sql = s_Sql & "ORDER BY item, secuencia"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Cargo las opciones
  While Not porstRecordset.EOF
    nContador = 0
    If porstRecordset("secuencia") = "0" Then
      outSeguridad.ListIndex = -1
      outSeguridad.AddItem Left(porstRecordset("detmdl"), InStr(porstRecordset("detmdl"), "&") - 1) & Mid(porstRecordset("detmdl"), InStr(porstRecordset("detmdl"), "&") + 1)
      n_Index = outSeguridad.ListCount - 1
      nContador = 1
    Else
      outSeguridad.ListIndex = n_Index
      outSeguridad.AddItem Left(porstRecordset("detmdl"), InStr(porstRecordset("detmdl"), "&") - 1) & Mid(porstRecordset("detmdl"), InStr(porstRecordset("detmdl"), "&") + 1)
      nContador = 1
    End If
    ' Redimensiono el arreglo de opciones
    If nContador = 1 Then
      ReDim Preserve a_Opciones(outSeguridad.ListCount)
      a_Opciones(outSeguridad.ListCount) = porstRecordset("codmdl")
    End If
    porstRecordset.MoveNext
    DoEvents
  Wend
  porstRecordset.Close
  
  ' Actualizo(activo, desactivo) las opciones
  For n_Index = 1 To outSeguridad.ListCount - 1
    If outSeguridad.Indent(n_Index) = 1 Then
      If outSeguridad.HasSubItems(n_Index) Then
        outSeguridad.PictureType(n_Index) = outOpen
      Else
        outSeguridad.RemoveItem n_Index
      End If
    Else
      outSeguridad.PictureType(n_Index) = outClosed
    End If
  Next n_Index

End Sub
Private Sub Empresas_Usuario(sUsuario As String)
  Dim nEmpresa As Integer, nPicture As Integer
  
  ' Obtengo las empresas por usuario
  s_Sql = "SELECT DISTINCTROW emp.codemp, seg.codusr "
  s_Sql = s_Sql & "FROM tgemp emp "
  s_Sql = s_Sql & "LEFT JOIN sgpms seg ON emp.codemp=seg.codemp AND seg.codsis='" & ps_CodSistema & "' AND seg.codusr='" & sUsuario & "' "
  s_Sql = s_Sql & "WHERE IFNULL(emp.codemp, '')<>'' "
  s_Sql = s_Sql & "ORDER BY codemp"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    
  nEmpresa = 0
  n_Index = nEmpresa
  ' Inicializo las empresas por usuario
  While Not porstRecordset.EOF
    n_Index = n_Index + 1
    nPicture = IIf(gdl_Funcion.aTexto(porstRecordset!codusr) = "", outClosed, outLeaf)
    outEmpresa.PictureType(n_Index) = nPicture
    nEmpresa = IIf((nPicture = outLeaf And nEmpresa = 0), n_Index, nEmpresa)
    porstRecordset.MoveNext
  Wend
  porstRecordset.Close
  ' Seleciono la empresa y cargo las opciones
  outEmpresa.ListIndex = nEmpresa
  Opciones_Usuario a_Empresa(outEmpresa.ListIndex), sUsuario
 
End Sub
Private Sub Opciones_Usuario(s_Empresa As String, s_Usuario As String)
  Dim nPicture As Integer
  
  ' Obtengo las opciones del usuario
  s_Sql = "SELECT DISTINCTROW opcion, orden, mnu.codmdl, mnu.detmdl, seg.codusr, "
  s_Sql = s_Sql & "RIGHT(IF(mnu.opcion='0', mnu.orden, mnu.opcion), 1) AS item, "
  s_Sql = s_Sql & "IF(mnu.opcion='0', mnu.opcion, mnu.orden) AS secuencia "
  s_Sql = s_Sql & "FROM sgmdl mnu "
  s_Sql = s_Sql & "LEFT JOIN sgpms seg ON mnu.codmdl=seg.codmdl "
  s_Sql = s_Sql & "AND seg.codemp='" & s_Empresa & "' AND mnu.codsis=seg.codsis AND seg.codusr='" & s_Usuario & "' "
  s_Sql = s_Sql & "WHERE mnu.codsis='" & ps_CodSistema & "' "
  s_Sql = s_Sql & "AND mnu.detmdl<>'-' "
  s_Sql = s_Sql & "ORDER BY item, secuencia"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)

  n_Index = 0
  ' Inicializo las opciones por usuario
  While Not porstRecordset.EOF
    n_Index = n_Index + IIf(porstRecordset("opcion") = "0", 1, 0)
    If porstRecordset("opcion") <> "0" Then
      n_Index = n_Index + 1
      nPicture = IIf(gdl_Funcion.aTexto(porstRecordset!codusr) = "", outClosed, outLeaf)
      outSeguridad.PictureType(n_Index) = nPicture
    End If
    porstRecordset.MoveNext
  Wend
  porstRecordset.Close
  ' Seleciono la opcion
  outSeguridad.ListIndex = 1
    
End Sub
Private Sub cmdAction_Click(Index As Integer)
  
  If Index = 2 Then
    Unload Me
  Else
  End If

End Sub
Private Sub cmdHelp_Click(Index As Integer)

  s_SqlHelp = ""
  ' Recupero la información
  s_Sql = gdl_Funcion.HelpTablas("usr", tdbHelp.Columns(0).DataField, ps_CodEmpresa, "")
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = frmCuadro(0).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = frmCuadro(0).Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 6730: Me.Width = 7360
  gdl_Procedure.CentraFormulario Me
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización de Opciones de Usuario"
  lblTitle = "Opciones del Sistema"
  
  n_IndexHelp = -1: s_SqlHelp = ""
' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "segurida": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "prelimin", "Imprimir", "cancelar")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Presentación Preliminar", "Imprimir", "Cancelación de " & lblTitle)
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(2).Cancel = True

  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  gdl_Procedure.EditText "PK", txtUsuario, ps_Usuario, "A", False, 10
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUsuario, "NU")
  
  ' Carga las empresa y opciones del sistema
  Empresa_Opcion
  Empresas_Usuario txtUsuario

 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Nombre")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codusr", "nomusr")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Entidad Bancaria", 2
  ']

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub outEmpresa_GotFocus()
  If txtUsuario = "" Then Beep: MsgBox "Debe Ingresar usuario del sistema", vbExclamation: txtUsuario.SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Usuario no es valido; Verificar", vbExclamation: txtUsuario.SetFocus: Exit Sub
End Sub
Private Sub outEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nListIndex As Integer

  If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    nListIndex = outEmpresa.ListIndex
    If (nListIndex > 0 And nListIndex < outEmpresa.ListCount - 1) Then
      If KeyCode = vbKeyDown Then nListIndex = nListIndex + 1 Else nListIndex = nListIndex - 1
    End If
    Opciones_Usuario a_Empresa(nListIndex), txtUsuario.Text
  End If

End Sub
Private Sub outEmpresa_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
    Opciones_Usuario a_Empresa(outEmpresa.ListIndex), txtUsuario.Text
  End If
End Sub
Private Sub outEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    Opciones_Usuario a_Empresa(outEmpresa.ListIndex), txtUsuario.Text
  End If
End Sub
Private Sub outEmpresa_PictureDblClick(ListIndex As Integer)
  Dim nPicture As Integer
     
  If outEmpresa.Indent(ListIndex) = 1 Then
    ' Coloco el puntero en normal
    gdl_Procedure.PunteroNormal
    ' Ubico en el indice actualizar
    outEmpresa.ListIndex = ListIndex
    nPicture = outEmpresa.PictureType(ListIndex)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_BDSystems)
  
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    ' Elimino la seguridad de la empresa
    If nPicture = outLeaf Then
      s_Sql = "DELETE FROM sgpms "
      s_Sql = s_Sql & "WHERE codemp='" & a_Empresa(ListIndex) & "' "
      s_Sql = s_Sql & "AND codusr='" & txtUsuario.Text & "' "
      s_Sql = s_Sql & "AND codsis='" & ps_CodSistema & "'"
      If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
    End If
    ' Actualizo la seguridad de la empresa
    If nPicture = outClosed Then
      s_Sql = "INSERT INTO sgpms(codusr, codemp, codmdl, codsis, usrcre, fyhcre) "
      s_Sql = s_Sql & "SELECT '" & txtUsuario.Text & "', '" & a_Empresa(ListIndex) & "', mnu.codmdl, '" & ps_CodSistema & "', "
      s_Sql = s_Sql & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
      s_Sql = s_Sql & "FROM sgmdl mnu "
      s_Sql = s_Sql & "WHERE mnu.codsis='" & ps_CodSistema & "' "
      s_Sql = s_Sql & "AND mnu.opcion<>'0' "
      s_Sql = s_Sql & "AND mnu.detmdl<>'-' "
      s_Sql = s_Sql & "ORDER BY mnu.opcion, mnu.orden, mnu.codmdl"
      If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
    End If
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    outEmpresa.PictureType(ListIndex) = IIf(nPicture = outLeaf, outClosed, outLeaf)
    ' Refresco las opciones del usuario
    Opciones_Usuario a_Empresa(ListIndex), txtUsuario.Text
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub outSeguridad_PictureClick(ListIndex As Integer)
  Dim nPicture As Integer
  
  If txtUsuario = "" Then Beep: MsgBox "Debe Ingresar usuario del sistema", vbExclamation: txtUsuario.SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Usuario no es valido; Verificar", vbExclamation: txtUsuario.SetFocus: Exit Sub
  If outEmpresa.ListIndex = 0 Then Beep: MsgBox "Debe selecionar una empresa", vbExclamation: Exit Sub
  If outSeguridad.Indent(ListIndex) > 1 Then
    ' Coloco el puntero en normal
    gdl_Procedure.PunteroNormal
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_BDSystems)
  
    nPicture = outSeguridad.PictureType(ListIndex)
    ' Elimino la seguridad de la empresa
    If nPicture = outLeaf Then
      s_Sql = "DELETE FROM sgpms "
      s_Sql = s_Sql & "WHERE codemp='" & a_Empresa(outEmpresa.ListIndex) & "' "
      s_Sql = s_Sql & "AND codusr='" & txtUsuario.Text & "' "
      s_Sql = s_Sql & "AND codsis='" & ps_CodSistema & "' "
      s_Sql = s_Sql & "AND codmdl='" & a_Opciones(ListIndex + 1) & "'"
    End If
    ' Actualizo la seguridad de la empresa
    If nPicture = outClosed Then
      s_Sql = "INSERT INTO sgpms(codusr, codemp, codmdl, codsis, usrcre, fyhcre) "
      s_Sql = s_Sql & "VALUES('" & txtUsuario.Text & "', '" & a_Empresa(outEmpresa.ListIndex) & "', '" & a_Opciones(ListIndex + 1) & "', "
      s_Sql = s_Sql & "'" & ps_CodSistema & "', '" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "')"
    End If
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    ' Activo o desactivo la opción
    outSeguridad.PictureType(ListIndex) = IIf(nPicture = outLeaf, outClosed, outLeaf)
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  txtUsuario.Text = tdbHelp.Columns(0).Value
  lblHelp(0) = tdbHelp.Columns(1).Value
  txtUsuario.SetFocus
  
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  ' Recupero la información ordenada
  s_Sql = gdl_Funcion.HelpTablas("usr", tdbHelp.Columns(ColIndex).DataField, ps_CodEmpresa, "")
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtUsuario_GotFocus()
  gdl_Procedure.MarcaGet txtUsuario
End Sub
Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    outSeguridad.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtUsuario_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUsuario, "NU")
  Empresas_Usuario txtUsuario
End Sub
