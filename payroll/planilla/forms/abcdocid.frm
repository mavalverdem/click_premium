VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcDocIdentidad 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7200
   Icon            =   "abcdocid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   7200
   Begin TabDlg.SSTab tabRegister 
      Height          =   2205
      Left            =   75
      TabIndex        =   22
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
      TabPicture(0)   =   "abcdocid.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDescripcion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtSiglas"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboSunat"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.ComboBox cboSunat 
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "abcdocid.frx":0028
         Left            =   1320
         List            =   "abcdocid.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtSiglas 
         Height          =   300
         Left            =   1335
         TabIndex        =   4
         Top             =   960
         Width           =   980
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         MaxLength       =   40
         TabIndex        =   2
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
         TabIndex        =   0
         Top             =   270
         Width           =   980
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   765
         Index           =   0
         Left            =   4515
         TabIndex        =   6
         Top             =   1005
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   1349
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
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
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
            Left            =   120
            TabIndex        =   8
            Top             =   480
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
         Caption         =   "Codigo Sunat :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1125
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
         TabIndex        =   23
         Top             =   315
         Width           =   1000
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Siglas-Abrev :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripci�n :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   5
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
         Picture         =   "abcdocid.frx":002C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   10
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
         Picture         =   "abcdocid.frx":0048
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
         TabIndex        =   11
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   12
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
         Picture         =   "abcdocid.frx":0064
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   14
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
         Picture         =   "abcdocid.frx":0080
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
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
         Picture         =   "abcdocid.frx":009C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   16
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
         Picture         =   "abcdocid.frx":00B8
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   2205
      Index           =   0
      Left            =   6435
      TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         Picture         =   "abcdocid.frx":00D4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   20
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
         Picture         =   "abcdocid.frx":00F0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   21
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
         Picture         =   "abcdocid.frx":010C
      End
   End
End
Attribute VB_Name = "fAbcDocIdentidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificaci�n de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String                ' Indice para bucle, y parametro de codigo
Private s_DocuIdentidad As String                       ' Codigo del registro
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acci�n
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creaci�n", IIf(Me.Tag = s_MdoData_Del, " - Eliminaci�n", IIf(Me.Tag = s_MdoData_Upd, " - Actualizaci�n", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
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
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!coddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!desdci.DefinedSize
    gdl_Procedure.EditText "AT", txtSiglas, "", Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!sigladci.DefinedSize
    gdl_Procedure.EditCombo "AT", cboSunat, -1, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
Else
    gdl_Procedure.EditText "PK", txtCodigo, fTablaSistema.dcaRegistro.Recordset!coddci, Me.Tag, True, fTablaSistema.dcaRegistro.Recordset!coddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fTablaSistema.dcaRegistro.Recordset!desdci), Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!desdci.DefinedSize
    gdl_Procedure.EditText "AT", txtSiglas, gdl_Funcion.aTexto(fTablaSistema.dcaRegistro.Recordset!sigladci), Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!sigladci.DefinedSize
    For n_Index = 0 To cboSunat.ListCount
      If cboSunat.List(n_Index) = gdl_Funcion.aTexto(fTablaSistema.dcaRegistro.Recordset!codsunat) Then Exit For
    Next n_Index
    n_Index = IIf(n_Index > cboSunat.ListCount, -1, n_Index)
    gdl_Procedure.EditCombo "AT", cboSunat, n_Index, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fTablaSistema.dcaRegistro.Recordset!estadodci = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fTablaSistema.dcaRegistro.Recordset!estadodci = s_Estado_Ina), Me.Tag, True
End If

End Sub
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
    s_DocuIdentidad = Trim$(txtCodigo)
    
    '[ Inicio la conexi�n a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("coddci")
    a_Valores = Array(s_DocuIdentidad)
    a_Tipos = Array(TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacci�n
    ' Elimino el registro
    If Not Records_Del("pldocidentidad", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacci�n
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fTablaSistema.dcaRegistro, fTablaSistema.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fTablaSistema.dcaRegistro.Recordset.EOF And fTablaSistema.dcaRegistro.Recordset.BOF) Or fTablaSistema.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fTablaSistema.dcaRegistro.Recordset.Find ("coddci >= '" & s_DocuIdentidad & "'")
      If fTablaSistema.dcaRegistro.Recordset.EOF Then fTablaSistema.dcaRegistro.Recordset.MoveLast
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
   Case 0: fTablaSistema.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fTablaSistema.dcaRegistro.Recordset.BOF Then fTablaSistema.dcaRegistro.Recordset.MovePrevious
           If fTablaSistema.dcaRegistro.Recordset.BOF Then fTablaSistema.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fTablaSistema.dcaRegistro.Recordset.EOF Then fTablaSistema.dcaRegistro.Recordset.MoveNext
           If fTablaSistema.dcaRegistro.Recordset.EOF Then fTablaSistema.dcaRegistro.Recordset.MoveLast
   Case 3: fTablaSistema.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1, s_Sunat As String * 100

  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripci�n " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If cboSunat.Text = "" Then Beep: MsgBox "Seleccione Codigo Sunat " & lblTitle, vbExclamation: cboSunat.SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, s_Estado_Act, s_Estado_Ina)
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_DocuIdentidad = txtCodigo
  s_Sunat = Trim(cboSunat.Text)
    
  ' Creo los arreglos para la actualizaci�n
  a_Campos = Array("coddci", "desdci", "sigladci", "codsunat", "estadodci", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(txtCodigo, Trim$(txtDescripcion), Trim$(txtSiglas), s_Sunat, s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("coddci")
  
  '[ Inicio la conexi�n a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacci�n
  ' Realizo el proceso de actualizaci�n de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("pldocidentidad", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("pldocidentidad", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacci�n
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fTablaSistema.dcaRegistro, fTablaSistema.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fTablaSistema.dcaRegistro.Recordset.Find ("coddci='" & s_DocuIdentidad & "'")
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
  s_TitleWindow = "Actualizaci�n Documento de Identidad"
  lblTitle = "Documento Identidad"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operaci�n del registro
  Me.Tag = fTablaSistema.Tag
  
  ' Configuro parametros de visualizaci�n del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y t�tulo del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "A�adir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualizaci�n del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y t�tulo del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualizaci�n
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Informaci�n de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Informaci�n de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fTablaSistema.dcaRegistro.Recordset.EOF Or fTablaSistema.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fTablaSistema.dcaRegistro.Recordset!coddci
  ' Cargo valores documento sunat
  For n_Index = 0 To 4
    cboSunat.AddItem Choose(n_Index + 1, "01 DOC. NACIONAL DE IDENTIDAD", "04 CARN� DE EXTRANJER�A", "07 PASAPORTE", "11 PARTIDA DE NACIMIENTO", "99 NINGUNO")
  Next n_Index

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
    txtSiglas.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtSiglas_GotFocus()
  gdl_Procedure.MarcaGet txtSiglas
End Sub
Private Sub txtSiglas_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    If optEstado(0).Enabled Then
      optEstado(0).SetFocus
    Else
      txtCodigo.SetFocus
    End If
    KeyAscii = 0
  End If

End Sub
