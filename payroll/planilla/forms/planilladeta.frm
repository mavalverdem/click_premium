VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fPlanillaDeta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7515
   Icon            =   "planilladeta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7515
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      _Version        =   65536
      _ExtentX        =   13256
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
         Left            =   6840
         TabIndex        =   3
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
         Picture         =   "planilladeta.frx":000C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6450
         TabIndex        =   4
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "planilladeta.frx":0028
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
         TabIndex        =   1
         Top             =   120
         Width           =   5850
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   4485
      Width           =   7515
      _Version        =   65536
      _ExtentX        =   13256
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
   Begin Threed.SSFrame frmCuadro 
      Height          =   3525
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   930
      Width           =   7425
      _Version        =   65536
      _ExtentX        =   13097
      _ExtentY        =   6218
      _StockProps     =   14
      Caption         =   " Parametros "
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
      Begin VB.ListBox lstConcepto 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   2790
         Index           =   1
         Left            =   4140
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   645
         Width           =   3210
      End
      Begin VB.ListBox lstConcepto 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   2790
         Index           =   0
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   645
         Width           =   3210
      End
      Begin Threed.SSPanel panToolBar 
         Height          =   2790
         Index           =   0
         Left            =   3345
         TabIndex        =   11
         Top             =   645
         Width           =   750
         _Version        =   65536
         _ExtentX        =   1323
         _ExtentY        =   4921
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
            TabIndex        =   12
            Top             =   15
            Width           =   720
            _Version        =   65536
            _ExtentX        =   1270
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Acción"
            ForeColor       =   16711680
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
         Begin Threed.SSCommand cmdMove 
            Height          =   360
            Index           =   0
            Left            =   150
            TabIndex        =   13
            Tag             =   "0"
            Top             =   540
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   12632256
            BevelWidth      =   0
            Outline         =   0   'False
            AutoSize        =   2
            Picture         =   "planilladeta.frx":0044
         End
         Begin Threed.SSCommand cmdMove 
            Height          =   360
            Index           =   1
            Left            =   150
            TabIndex        =   14
            Tag             =   "0"
            Top             =   960
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   12632256
            BevelWidth      =   0
            Outline         =   0   'False
            AutoSize        =   2
            Picture         =   "planilladeta.frx":0060
         End
         Begin Threed.SSCommand cmdMove 
            Height          =   360
            Index           =   2
            Left            =   150
            TabIndex        =   15
            Tag             =   "0"
            Top             =   1680
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   12632256
            BevelWidth      =   0
            Outline         =   0   'False
            AutoSize        =   2
            Picture         =   "planilladeta.frx":007C
         End
         Begin Threed.SSCommand cmdMove 
            Height          =   360
            Index           =   3
            Left            =   150
            TabIndex        =   16
            Tag             =   "0"
            Top             =   2115
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   12632256
            BevelWidth      =   0
            Outline         =   0   'False
            AutoSize        =   2
            Picture         =   "planilladeta.frx":0098
         End
      End
      Begin Threed.SSOption optConcepto 
         Height          =   195
         Index           =   0
         Left            =   1035
         TabIndex        =   17
         Top             =   345
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Ingreso"
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
         Font3D          =   3
      End
      Begin Threed.SSOption optConcepto 
         Height          =   195
         Index           =   1
         Left            =   1950
         TabIndex        =   18
         Top             =   345
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Descuento"
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
         Font3D          =   3
      End
      Begin Threed.SSOption optConcepto 
         Height          =   195
         Index           =   2
         Left            =   3165
         TabIndex        =   19
         Top             =   345
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Aporte"
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
         Font3D          =   3
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   0
         Left            =   945
         Shape           =   4  'Rounded Rectangle
         Top             =   270
         Width           =   3105
      End
      Begin VB.Label lblDato 
         Caption         =   "Contenido Grupo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   4140
         TabIndex        =   7
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label lblDato 
         Caption         =   "Concepto :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   420
         Width           =   1005
      End
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
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
      Left            =   255
      TabIndex        =   8
      Top             =   615
      Width           =   375
   End
End
Attribute VB_Name = "fPlanillaDeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private s_Tabla As String                               ' Tabla temporal del detalle
Private n_Fila As Integer, n_Columna As Integer         ' Fila y columna de la planilla
'[
Private Sub RecuperaRegistros(ByVal sExpresion As String)

  ' Inicializo la lista de conceptos
  lstConcepto(0).Clear
  s_Sql = "SELECT cxp.codcpc, cpc.descpc "
  s_Sql = s_Sql & "FROM plconceplanilla cxp "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON cxp.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE cxp.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND cpc.tipocpc='" & sExpresion & "' "
  s_Sql = s_Sql & "AND NOT EXISTS(SELECT * FROM " & s_Tabla & " tmp "
  s_Sql = s_Sql & "WHERE tmp.codcls=cxp.codcls "
  s_Sql = s_Sql & "AND tmp.codcpc=cxp.codcpc "
  s_Sql = s_Sql & "AND tmp.fila=" & n_Fila & " "
  s_Sql = s_Sql & "AND tmp.columna=" & n_Columna & ") "
  s_Sql = s_Sql & "ORDER BY codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    While Not porstRecordset.EOF
      lstConcepto(0).AddItem porstRecordset!codcpc & " - " & porstRecordset!descpc
      porstRecordset.MoveNext
    Wend
  End If

  ' Inicializo la lista de conceptos por grupo
  lstConcepto(1).Clear
  s_Sql = "SELECT tmp.codcpc, cpc.descpc "
  s_Sql = s_Sql & "FROM " & s_Tabla & " tmp "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON tmp.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE tmp.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND cpc.tipocpc='" & sExpresion & "' "
  s_Sql = s_Sql & "AND tmp.fila=" & n_Fila & " "
  s_Sql = s_Sql & "AND tmp.columna=" & n_Columna & " "
  s_Sql = s_Sql & "ORDER BY codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    While Not porstRecordset.EOF
      lstConcepto(1).AddItem porstRecordset!codcpc & " - " & porstRecordset!descpc
      porstRecordset.MoveNext
    Wend
  End If

End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdMove_Click(Index As Integer)
  Dim nListIndex As Integer, nListCount As Integer
  
  If Index = 0 And lstConcepto(0).ListCount >= 1 Then
    nListIndex = lstConcepto(0).ListIndex
    If nListIndex = -1 Then Beep: MsgBox "Debe seleccionar concepto de agrupación", vbExclamation: lstConcepto(0).SetFocus: Exit Sub
    nListCount = lstConcepto(0).ListCount - 2
    lstConcepto(1).AddItem lstConcepto(0).Text
    lstConcepto(0).RemoveItem nListIndex
    nListIndex = IIf(nListCount >= nListIndex, nListIndex, nListCount)
    lstConcepto(0).ListIndex = nListIndex
  ElseIf Index = 1 And lstConcepto(0).ListCount >= 1 Then
    nListCount = lstConcepto(0).ListCount - 1
    For nListIndex = 0 To nListCount
      lstConcepto(0).ListIndex = nListIndex
      lstConcepto(1).AddItem lstConcepto(0).Text
    Next nListIndex
    lstConcepto(0).Clear
  ElseIf Index = 2 And lstConcepto(1).ListCount >= 1 Then
    nListIndex = lstConcepto(1).ListIndex
    If nListIndex = -1 Then Beep: MsgBox "Debe seleccionar concepto a desagrupar", vbExclamation: lstConcepto(1).SetFocus: Exit Sub
    nListCount = lstConcepto(1).ListCount - 2
    lstConcepto(0).AddItem lstConcepto(1).Text
    lstConcepto(1).RemoveItem nListIndex
    nListIndex = IIf(nListCount >= nListIndex, nListIndex, nListCount)
    lstConcepto(1).ListIndex = nListIndex
  ElseIf Index = 3 And lstConcepto(1).ListCount >= 1 Then
    nListCount = lstConcepto(1).ListCount - 1
    For nListIndex = 0 To nListCount
      lstConcepto(1).ListIndex = nListIndex
      lstConcepto(0).AddItem lstConcepto(1).Text
    Next nListIndex
    lstConcepto(1).Clear
  End If

End Sub
Private Sub cmdUpdate_Click()
  Dim sConcepto As String, sTipoConcepto As String
  Dim nLongitud As Integer
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  sTipoConcepto = IIf(optConcepto(0).Value, "0", IIf(optConcepto(1).Value, "1", "2"))
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpll", "fila", "columna", "codcpc", "tipocpc", "usrcre", "fyhcre")
  a_Valores = Array(ps_ClsPlanilla, Trim(fAbcPlanillaGnral.txtCodigo.Text), n_Fila, n_Columna, "", sTipoConcepto, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpll", "fila", "columna")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  gdl_Conexion.IniciaTransaccion    'Inicia transacción
  ' Elimino los registros anteriores
  If Not Records_Del(s_Tabla, a_Where, a_Valores, a_Tipos) Then GoTo Error
  ' Realizo el proceso de actualización de los registros
  For n_Index = 0 To lstConcepto(1).ListCount - 1
    lstConcepto(1).ListIndex = n_Index
    nLongitud = InStr(lstConcepto(1).Text, "-") - 1
    sConcepto = Trim(Left(lstConcepto(1).Text, nLongitud))
    a_Valores = Array(ps_ClsPlanilla, Trim(fAbcPlanillaGnral.txtCodigo.Text), n_Fila, n_Columna, sConcepto, sTipoConcepto, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
    If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Next n_Index
  gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
  
  ' Cierro el formulario
  Unload Me
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
  Dim sTipoCpc As String
  
  'Establece posición y titulo del formulario
  Me.Height = 5475: Me.Width = 7605
  Me.Left = 2800: Me.Top = 2600

  ' Titulo del formulario y panel
  s_TitleWindow = "Agrupación de Detalle Planilla"
  lblTitle = "Detalle de Planilla"
  s_Tabla = fAbcPlanillaGnral.txtCodigo.Tag

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "siguient", "ultimo", "anterior", "primero")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Asignar ", "Todos ", "Desasignar ", "Todos ") & "Conceptos"
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento

  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True

  ' Carga los datos en el formulario
  n_Fila = CLng(Left(fAbcPlanillaGnral.tdbDetalle.Columns(12).Text, 2))
  n_Columna = CLng(Mid(fAbcPlanillaGnral.tdbDetalle.Columns(12).Text, 4))
  lblTexto.Caption = Trim(fAbcPlanillaGnral.txtDescripcion.Text) & " - " & Trim(fAbcPlanillaGnral.tdbDetalle.Columns(0).Text) & "/" & Trim(fAbcPlanillaGnral.tdbDetalle.Columns(2).Text) & "  " & Trim(fAbcPlanillaGnral.tdbDetalle.Columns(5).Text) & " "
  sTipoCpc = "0"
  s_Sql = "SELECT DISTINCT IFNULL(tipocpc, '0') AS tipocpc "
  s_Sql = s_Sql & "FROM " & s_Tabla & " "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpll='" & fAbcPlanillaGnral.txtCodigo.Text & "' "
  s_Sql = s_Sql & "AND fila=" & n_Fila & " "
  s_Sql = s_Sql & "AND columna=" & n_Columna
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not ((porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount = 0) Then sTipoCpc = porstRecordset!tipocpc
  gdl_Procedure.EditOptionCheck "AT", optConcepto(0), (sTipoCpc = "0"), s_MdoData_Ins, True
  gdl_Procedure.EditOptionCheck "AT", optConcepto(1), (sTipoCpc = "1"), s_MdoData_Ins, True
  gdl_Procedure.EditOptionCheck "AT", optConcepto(2), (sTipoCpc = "2"), s_MdoData_Ins, True
  ']

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub optConcepto_Click(Index As Integer, Value As Integer)
  ' Verifico la opcion seleccionada
  RecuperaRegistros Index
End Sub
