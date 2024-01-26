VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcCentroCosto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7200
   Icon            =   "abccecos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   7200
   Begin TabDlg.SSTab tabRegister 
      Height          =   2205
      Left            =   75
      TabIndex        =   21
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   3889
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
      TabPicture(0)   =   "abccecos.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmCuadro(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDescripcion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         MaxLength       =   40
         TabIndex        =   3
         Top             =   615
         Width           =   4680
      End
      Begin VB.TextBox txtCodigo 
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
         Height          =   300
         Left            =   1340
         TabIndex        =   1
         Top             =   270
         Width           =   1125
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   645
         Index           =   0
         Left            =   2955
         TabIndex        =   4
         Top             =   1005
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   " Estado "
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
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   5
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Activo"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   1
            Left            =   1635
            TabIndex        =   6
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Inactivo"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripci�n :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "C�digo :"
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
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   1000
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Left            =   6450
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
         Picture         =   "abccecos.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   9
         Top             =   75
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abccecos.frx":0044
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
         Left            =   675
         TabIndex        =   10
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   2895
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Left            =   4695
         TabIndex        =   12
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
         Picture         =   "abccecos.frx":0060
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   13
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
         Picture         =   "abccecos.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   14
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
         Picture         =   "abccecos.frx":0098
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   15
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
         Picture         =   "abccecos.frx":00B4
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   2205
      Index           =   0
      Left            =   6435
      TabIndex        =   16
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   3889
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
         TabIndex        =   17
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edici�n"
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
         TabIndex        =   18
         Tag             =   "0"
         Top             =   420
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
         Picture         =   "abccecos.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Tag             =   "0"
         Top             =   1050
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
         Picture         =   "abccecos.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Tag             =   "0"
         Top             =   1650
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
         Picture         =   "abccecos.frx":0108
      End
   End
End
Attribute VB_Name = "fAbcCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificaci�n de existencia de Registros
Private i As Byte, s_ParCodigo As String                ' Indice para bucle, y parametro de codigo
Private s_CentroCosto As String, s_NivelCCo As String   ' Codigo del registro y niveles de centro de costo
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acci�n
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creaci�n", IIf(Me.Tag = s_MdoData_Del, " - Eliminaci�n", IIf(Me.Tag = s_MdoData_Upd, " - Actualizaci�n", " - Consulta")))
  For i = 0 To 3: cmdMove(i).Visible = (Me.Tag = s_MdoData_Vis): Next i
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)

End Sub
Sub ShowScreen()
    
' Presenta Botones y Controles
EnabledBotons
' Presenta datos en pantalla de acuerdo al modo Seleccionado
If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fCentroCosto.dcaRegistro.Recordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fCentroCosto.dcaRegistro.Recordset!detcco.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
Else
    gdl_Procedure.EditText "PK", txtCodigo, fCentroCosto.dcaRegistro.Recordset!codcco, Me.Tag, True, fCentroCosto.dcaRegistro.Recordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, fCentroCosto.dcaRegistro.Recordset!detcco, Me.Tag, False, fCentroCosto.dcaRegistro.Recordset!detcco.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fCentroCosto.dcaRegistro.Recordset!estcco = "A"), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fCentroCosto.dcaRegistro.Recordset!estcco = "I"), Me.Tag, True
End If

End Sub
Function ValidaCenCos(ByVal s_CenCosto As String) As Boolean
  Dim n_Longitud As Integer, n_Posicion As Integer
  Dim n_Index As Integer, n_Registro As Long
  Dim s_Where As String, s_ItemCenCos As String, s_Mensage As String
  
  ValidaCenCos = True
  n_Longitud = Len(Trim$(s_CenCosto))
  s_Sql = "SELECT COUNT(*) AS nRegistro"
  s_Sql = s_Sql & " FROM cocco"
  If Me.Tag = s_MdoData_Ins Then
    If n_Longitud <> Left(s_NivelCCo, 1) Then
      ' Posicion de configuracion de nivel
      n_Posicion = InStr(s_NivelCCo, n_Longitud)
      s_Mensage = "No configurado Nivel '" & Left(s_CenCosto, n_Longitud) & "' de Centros de Costos, Verificar"
      If Not (n_Posicion > 0) Then GoTo ErrorValid
      ' Niveles del centro de costo
      For n_Index = (n_Posicion - 1) To 1 Step -1
        s_ItemCenCos = Left(s_CenCosto, Mid(s_NivelCCo, n_Index, 1))
        s_Where = s_Sql & " WHERE codcco='" & s_ItemCenCos & "'"
        Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Where)
        s_Mensage = "No Existen Nivel '" & s_ItemCenCos & "' de Centros de Costos, Verificar"
        n_Registro = porstRecordset!nRegistro
        ' Cierro el recordset
        porstRecordset.Close: Set porstRecordset = Nothing
        If n_Registro = 0 Then GoTo ErrorValid
      Next n_Index
    End If
  ElseIf Me.Tag = s_MdoData_Del Then
    s_Sql = s_Sql & " WHERE LEFT(codcco, " & n_Longitud & ")='" & s_CenCosto & "'"
    s_Sql = s_Sql & " AND LENGTH(codcco)>" & n_Longitud
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    s_Mensage = "Existen Centros de Costos Relacionados, No podra Eliminar"
    n_Registro = porstRecordset!nRegistro
    ' Cierro el recordset
    porstRecordset.Close: Set porstRecordset = Nothing
    If n_Registro > 0 Then GoTo ErrorValid
  End If
  Exit Function

ErrorValid:
  MsgBox s_Mensage, vbExclamation, "Advertencia"
  ValidaCenCos = False
  
End Function
']
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   txtDescripcion.SetFocus
  End If
  If Index <> 1 Then Exit Sub
    
  Beep
  If MsgBox("� Est�s Seguro de Eliminar el " & lblTitle & " '" & Trim$(txtDescripcion) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_CentroCosto = Trim$(txtCodigo)
    ' Valido niveles de centro de costo
    If Not ValidaCenCos(s_CentroCosto) Then GoTo Finalizar
    
    '[ Inicio la conexi�n a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DaBasCon)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcco")
    a_Valores = Array(s_CentroCosto)
    a_Tipos = Array(TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacci�n
    ' Elimino el registro
    If Not Records_Del("cocco", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacci�n
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fCentroCosto.dcaRegistro, fCentroCosto.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fCentroCosto.dcaRegistro.Recordset.EOF And fCentroCosto.dcaRegistro.Recordset.BOF) Or fCentroCosto.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fCentroCosto.dcaRegistro.Recordset.Find ("codcco >= '" & s_CentroCosto & "'")
      If fCentroCosto.dcaRegistro.Recordset.EOF Then fCentroCosto.dcaRegistro.Recordset.MoveLast
    Else
      Unload Me
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexi�n a la base de datos ]
  Set gdl_Conexion = Nothing
  If Not l_ExistRecord Then cmdCancel_Click
    
End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fCentroCosto.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fCentroCosto.dcaRegistro.Recordset.BOF Then fCentroCosto.dcaRegistro.Recordset.MovePrevious
           If fCentroCosto.dcaRegistro.Recordset.BOF Then fCentroCosto.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fCentroCosto.dcaRegistro.Recordset.EOF Then fCentroCosto.dcaRegistro.Recordset.MoveNext
           If fCentroCosto.dcaRegistro.Recordset.EOF Then fCentroCosto.dcaRegistro.Recordset.MoveLast
   Case 3: fCentroCosto.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1

  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripci�n " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, "A", "I")
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_CentroCosto = txtCodigo
  ' Valido el nivel del registro
  If Me.Tag = s_MdoData_Ins Then
    If Not ValidaCenCos(s_CentroCosto) Then txtCodigo.SetFocus: GoTo Finalizar
  End If
    
  ' Creo los arreglos para la actualizaci�n
  a_Campos = Array("codcco", "detcco", "estcco", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(txtCodigo, Trim$(txtDescripcion), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcco")
  
  '[ Inicio la conexi�n a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DaBasCon)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacci�n
  ' Realizo el proceso de actualizaci�n de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("cocco", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("cocco", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacci�n
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fCentroCosto.dcaRegistro, fCentroCosto.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fCentroCosto.dcaRegistro.Recordset.Find ("codcco='" & s_CentroCosto & "'")
  ' si es actualizaci�n pasa al modo visualizaci�n
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtCodigo.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexi�n a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminaci�n
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posici�n y titulo del formulario
  Me.Height = 3885: Me.Width = 7290
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualizaci�n Centro de Costos"
  lblTitle = "Centro de Costo"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operaci�n del registro
  Me.Tag = fCentroCosto.Tag
  s_NivelCCo = fCentroCosto.ps_NivelCCo
  
  ' Configuro parametros de visualizaci�n del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y t�tulo del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For i = 0 To 2
    aElemento(i, 1) = Choose(i + 1, "anadir", "borrar", "modifica")
    aElemento(i, 2) = Choose(i + 1, "A�adir ", "Eliminar ", "Modificar ") & lblTitle
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualizaci�n del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y t�tulo del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For i = 0 To 3
    aElemento(i, 1) = Choose(i + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(i, 2) = Choose(i + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next i
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualizaci�n
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Informaci�n de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Informaci�n de " & lblTitle
  cmdCancel.Cancel = True
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fCentroCosto.dcaRegistro.Recordset.EOF Or fCentroCosto.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fCentroCosto.dcaRegistro.Recordset!codcco
  ' Carga los datos en el formulario
  ShowScreen
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub txtCodigo_GotFocus()
  gdl_Procedure.MarcaGet txtCodigo
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If txtCodigo = "" Then
      Beep
      MsgBox "Debe Ingresar el C�digo del " & lblTitle, vbExclamation
      txtCodigo.SetFocus
    Else
      txtDescripcion.SetFocus
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txtDescripcion_GotFocus()
  gdl_Procedure.MarcaGet txtDescripcion
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If optEstado(0).Enabled Then
      optEstado(0).SetFocus
    Else
      txtCodigo.SetFocus
    End If
    KeyAscii = 0
  End If

End Sub