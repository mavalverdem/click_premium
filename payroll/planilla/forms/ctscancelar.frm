VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form fCtsCancelacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3225
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6645
   Icon            =   "ctscancelar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   6645
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6645
      _Version        =   65536
      _ExtentX        =   11721
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
         Index           =   0
         Left            =   5940
         TabIndex        =   9
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
         Picture         =   "ctscancelar.frx":000C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   13
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
         Left            =   675
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
         Left            =   1485
         TabIndex        =   10
         Top             =   120
         Width           =   3990
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   2055
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   600
      Width           =   6525
      _Version        =   65536
      _ExtentX        =   11509
      _ExtentY        =   3625
      _StockProps     =   14
      Caption         =   " Datos de Cancelación "
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
      Font3D          =   1
      ShadowStyle     =   1
      Begin VB.TextBox txtPeriodo 
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
         Height          =   280
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   2
         Top             =   435
         Width           =   1050
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   2145
         TabIndex        =   12
         Top             =   435
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   795
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmdProceso 
         Height          =   315
         Left            =   2235
         TabIndex        =   6
         Top             =   1455
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Procesar"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
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
         Left            =   2550
         TabIndex        =   3
         Top             =   480
         Width           =   195
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   795
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   945
         Index           =   1
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   270
         Width           =   6435
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   2715
      Width           =   6645
      _Version        =   65536
      _ExtentX        =   11721
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
      Left            =   1845
      TabIndex        =   11
      Top             =   720
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
Attribute VB_Name = "fCtsCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Private Sub cmdCancel_Click(Index As Integer)
  Unload Me
End Sub
Private Sub cmdHelp_Click(Index As Integer)

  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Periodo cts no cancelados
    tdbHelp.Columns(0).DataField = "pdocts": tdbHelp.Columns(1).DataField = "descricts"
    tdbHelp.Caption = "Periodo de Provisión CTS"
    s_Sql = gdl_Funcion.HelpTablas("cxe", "pdocts", IIf(ribParametro(0).Value, s_Estado_Act, s_Estado_Blq) & ps_ClsPlanilla, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = 300 + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = 1650
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdProceso_Click()
  Dim s_OldMessage As String, sFecha As String
  Dim sFechaHora As String, sProceso As String
  Dim nRegistros As Long

  If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Periodo de Liquidación de CTS", vbExclamation: txtPeriodo.SetFocus: Exit Sub
  If (lblHelp(0) = "" Or lblHelp(0) = "???") Then Beep: MsgBox "Periodo de Liquidación de CTS no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
  If mskFecha.ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha de Cancelación", vbExclamation: mskFecha.SetFocus: Exit Sub
  If mskFecha.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecha, 1900) Then mskFecha.SetFocus: Exit Sub
  End If
  ' Validación de cancelación
  If ribParametro(0).Value Then
    s_Sql = "SELECT fechafin FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND estadosub='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "ORDER BY subcts desc"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    sFecha = Format(porstRecordset!fechafin, "yyyymmdd")
    If Not (Format(mskFecha, "yyyymmdd") >= sFecha) Then Beep: MsgBox "Fecha de Cancelación debe ser mayor o Igual que fecha final del periodo", vbExclamation: mskFecha.SetFocus: Exit Sub
        
    ' Verifico no existan sub periodos provisionados
    s_Sql = "SELECT COUNT(*) AS registros FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND estadosub='" & s_Estado_Act & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If porstRecordset!registros = 0 Then Beep: MsgBox "No existe sub periodos procesados de  " & lblTitle.Caption, vbCritical: Exit Sub
  End If
  Beep
  If MsgBox("¿ Estás Seguro de " & IIf(ribParametro(0).Value, "Procesar", "Eliminar") & " Cancelación Sub Periodos de '" & lblHelp(0).Caption & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    fMenu.panPercent.Visible = True
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    sFechaHora = Format(Now, s_FmtFeHoMysql_0)
    sProceso = "cancelacts"
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To o_PvsComxTieSer.tdbRegistro.SelBookmarks.Count - 1
      o_PvsComxTieSer.tdbRegistro.Bookmark = o_PvsComxTieSer.tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, sProceso, o_PvsComxTieSer.tdbRegistro.Columns(0).Text, ps_Usuario, sFechaHora, "A"
    Next n_Index
    ' Incremento el porcentaje del proceso
    fMenu.panPercent.FloodPercent = 25
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    
    ' Actualizo (cancelo) los movimientos de cts
    s_Sql = "UPDATE plctsmovimiento mov, rangoimpresion rng "
    If ribParametro(0).Value Then
      s_Sql = s_Sql & "SET mov.estadomov='" & s_Estado_Blq & "', "
      s_Sql = s_Sql & "mov.fechacan='" & Format(mskFecha.Text, s_FmtFechMysql_0) & "', "
    Else
      s_Sql = s_Sql & "SET mov.estadomov='" & s_Estado_Act & "', "
      s_Sql = s_Sql & "mov.fechacan=Null, "
    End If
    s_Sql = s_Sql & "mov.usrmdf='" & ps_Usuario & "', "
    s_Sql = s_Sql & "mov.fyhmdf='" & sFechaHora & "' "
    s_Sql = s_Sql & "WHERE mov.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND mov.pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND mov.estadomov='" & IIf(ribParametro(0).Value, s_Estado_Act, s_Estado_Blq) & "' "
    s_Sql = s_Sql & "AND RTRIM(mov.codpsn)=RTRIM(rng.valor) "
    s_Sql = s_Sql & "AND rng.proceso='" & sProceso & "' "
    s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND rng.fyhcre='" & sFechaHora & "' "
    gdl_Conexion.Execucion s_Sql, Modifica
    ' Incremento el porcentaje del proceso
    fMenu.panPercent.FloodPercent = 50
    
    ' Actualizo el sub periodo de cts
    s_Sql = "UPDATE plctsperiodosub sub "
    If ribParametro(0).Value Then
      s_Sql = s_Sql & "SET sub.estadosub='" & s_Estado_Blq & "', "
      s_Sql = s_Sql & "sub.fechacan='" & Format(mskFecha.Text, s_FmtFechMysql_0) & "', "
    Else
      s_Sql = s_Sql & "SET sub.estadosub='" & s_Estado_Act & "', "
      s_Sql = s_Sql & "sub.fechacan=Null, "
    End If
    s_Sql = s_Sql & "sub.usrmdf='" & ps_Usuario & "', "
    s_Sql = s_Sql & "sub.fyhmdf='" & sFechaHora & "' "
    s_Sql = s_Sql & "WHERE sub.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND sub.pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND sub.estadosub='" & IIf(ribParametro(0).Value, s_Estado_Act, s_Estado_Blq) & "' "
    s_Sql = s_Sql & "AND EXISTS(SELECT * "
    s_Sql = s_Sql & "FROM plctsmovimiento mov "
    s_Sql = s_Sql & "WHERE mov.codcls=sub.codcls "
    s_Sql = s_Sql & "AND mov.pdocts=sub.pdocts "
    s_Sql = s_Sql & "AND mov.subcts=sub.subcts "
    s_Sql = s_Sql & "AND mov.estadomov='" & IIf(ribParametro(0).Value, s_Estado_Blq, s_Estado_Act) & "')"
    gdl_Conexion.Execucion s_Sql, Modifica
    ' Incremento el porcentaje del proceso
    fMenu.panPercent.FloodPercent = 75
    
    ' Verifico no existan sub periodos pendientes o provisionados
    s_Sql = "SELECT COUNT(*) AS registros FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND estadosub<>'" & s_Estado_Blq & "'"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    nRegistros = CLng(porstRecordset!registros)
    porstRecordset.Close
    
    ' Actualizo el periodo de provisión cts
    s_Sql = "UPDATE plctsperiodo "
    s_Sql = s_Sql & "SET estadocts='" & IIf(nRegistros = 0, s_Estado_Blq, s_Estado_Act) & "', "
    s_Sql = s_Sql & "usrmdf='" & ps_Usuario & "', "
    s_Sql = s_Sql & "fyhmdf='" & sFechaHora & "' "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND estadocts='" & IIf(nRegistros = 0, s_Estado_Act, s_Estado_Blq) & "'"
    gdl_Conexion.Execucion s_Sql, Modifica
    ' Incremento el porcentaje del proceso
    fMenu.panPercent.FloodPercent = 100
    
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    MsgBox "Se actualizo exitosamente " & lblTitle, vbInformation
  End If
  GoTo Finalizar
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Elimino el rango de impresion
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, sProceso, "", ps_Usuario, sFechaHora, "E"
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 3700: Me.Width = 6730
  Me.Left = 4580: Me.Top = 1300
  
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  ' Titulo del formulario y panel
  s_TitleWindow = "Cancelación de Liquidación de CTS"
  lblTitle = "Liquidación CTS"
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' paramteros de proceso
  For n_Index = 0 To 1
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = Choose(n_Index + 1, "Procesar ", "Eliminar ") & lblTitle.Caption
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "pagar", "borrar") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "proceso": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "cancelar"
  aElemento(0, 2) = "Cancelar Proceso de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdCancel, aElemento
  cmdProceso.ToolTipText = "Genera Proceso de " & lblTitle
  cmdCancel(0).Cancel = True
  
 '[ Configuración el control de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codpll", "despll")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Entidad Pensiones", 2
  ']
  ' Carga los datos en el formulario
  gdl_Procedure.EditText "AT", txtPeriodo, "", s_MdoData_Ins, False, 6
  gdl_Procedure.EditMask "AT", mskFecha, "", s_MdoData_Ins, True, "##/##/####"
  ribParametro(0).Value = True
  ']
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub mskFecha_GotFocus()
  gdl_Procedure.MarcaGet mskFecha
End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecha_Validate(Cancel As Boolean)
  If mskFecha.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecha, 1900
  End If
End Sub

Private Sub ribParametro_Click(Index As Integer, Value As Integer)
  txtPeriodo.Text = "": lblHelp(0).Caption = ""
  gdl_Procedure.EditMask "AT", mskFecha, "", s_MdoData_Ins, True, "##/##/####"
End Sub

Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Formato de planilla
    txtPeriodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Proceso de calculo
    s_Sql = gdl_Funcion.HelpTablas("cxe", tdbHelp.Columns(ColIndex).DataField, s_Estado_Act & ps_ClsPlanilla, "")
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
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus()
  Dim sFecha As String
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, "EC")
  sFecha = Format(Now, s_FormatoFecha)
  If Not (lblHelp(0) = "" Or lblHelp(0) = "???") Then
    s_Sql = "SELECT fechafin FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND estadosub='" & IIf(ribParametro(0).Value, s_Estado_Act, s_Estado_Blq) & "' "
    s_Sql = s_Sql & "ORDER BY subcts desc"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If porstRecordset.RecordCount >= 1 Then
      sFecha = Format(porstRecordset!fechafin, s_FormatoFecha)
    End If
    porstRecordset.Close
  End If
  gdl_Procedure.EditMask "AT", mskFecha, sFecha, s_MdoData_Ins, True, "##/##/####"
End Sub

