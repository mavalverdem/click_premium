VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fAbcSunat 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   7185
   Begin TabDlg.SSTab tabRegister 
      Height          =   2400
      Left            =   75
      TabIndex        =   0
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4233
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
      TabPicture(0)   =   "abcsunat.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmCuadro(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbtipo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDescripcion"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
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
         MaxLength       =   4
         TabIndex        =   3
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         MaxLength       =   40
         TabIndex        =   2
         Top             =   615
         Width           =   3360
      End
      Begin VB.ComboBox cmbtipo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "abcsunat.frx":001C
         Left            =   1320
         List            =   "abcsunat.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1280
         Width           =   2055
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1065
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Top             =   960
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   1879
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
            TabIndex        =   5
            Top             =   480
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
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
            Left            =   1080
            TabIndex        =   6
            Top             =   480
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
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
         Caption         =   "Código :"
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
         TabIndex        =   9
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo :"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   520
         TabIndex        =   7
         Top             =   1275
         Width           =   675
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
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
         TabIndex        =   11
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
         Picture         =   "abcsunat.frx":0020
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   12
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
         Picture         =   "abcsunat.frx":003C
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
         TabIndex        =   13
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   2400
      Index           =   0
      Left            =   6435
      TabIndex        =   14
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   4233
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
         TabIndex        =   15
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edición"
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
         TabIndex        =   16
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
         Picture         =   "abcsunat.frx":0058
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   17
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
         Picture         =   "abcsunat.frx":0074
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   18
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
         Picture         =   "abcsunat.frx":0090
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   19
      Top             =   3045
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
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
         TabIndex        =   20
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
         Picture         =   "abcsunat.frx":00AC
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   21
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
         Picture         =   "abcsunat.frx":00C8
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   22
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
         Picture         =   "abcsunat.frx":00E4
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   23
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
         Picture         =   "abcsunat.frx":0100
      End
      Begin MSAdodcLib.Adodc dcaHelp 
         Height          =   330
         Left            =   0
         Top             =   0
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
End
Attribute VB_Name = "fAbcSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private i As Byte, s_ParCodigo As String                ' Indice para bucle, y parametro de codigo
Private s_Sunat As String                               ' Codigo del registro
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private n_Index As Integer                              ' Indice para bucle
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
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
    
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!codcon.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!descon.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
        
    gdl_Procedure.EditCombo "AT", cmbtipo, -1, Me.Tag, False
    
Else
    gdl_Procedure.EditText "PK", txtCodigo, fTablaSistema.dcaRegistro.Recordset!codcon, Me.Tag, True, fTablaSistema.dcaRegistro.Recordset!codcon.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fTablaSistema.dcaRegistro.Recordset!descon), Me.Tag, False, fTablaSistema.dcaRegistro.Recordset!descon.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fTablaSistema.dcaRegistro.Recordset!estadocon = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fTablaSistema.dcaRegistro.Recordset!estadocon = s_Estado_Ina), Me.Tag, True
    
    
   cmbtipo.ForeColor = &HC00000
   If Me.Tag = "C" Then
      cmbtipo.Locked = False
      cmbtipo.BackColor = &H80000005
      Else
      cmbtipo.Locked = True
      cmbtipo.BackColor = &HC7D8E0
   End If
    
   cmbtipo.Text = fTablaSistema.dcaRegistro.Recordset!Tipcon
    
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
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim$(txtDescripcion) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Sunat = Trim$(txtCodigo)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcon")
    a_Valores = Array(s_Sunat)
    a_Tipos = Array(TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plconcesunat", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fTablaSistema.dcaRegistro, fTablaSistema.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fTablaSistema.dcaRegistro.Recordset.EOF And fTablaSistema.dcaRegistro.Recordset.BOF) Or fTablaSistema.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fTablaSistema.dcaRegistro.Recordset.Find ("codcon >= '" & s_Sunat & "'")
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
  '[ Finalizo la conexión a la base de datos ]
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
  Dim s_Estado As String * 1, s_Formato As String * 1, s_Tipo As String * 100

  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If cmbtipo = "" Then Beep: MsgBox "Seleccione Tipo " & lblTitle, vbExclamation: cmbtipo.SetFocus: Exit Sub
 
  
  
  s_Estado = IIf(optEstado(0).Value, s_Estado_Act, s_Estado_Ina)
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Sunat = txtCodigo.Text
  
  s_Tipo = cmbtipo.Text
      
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcon", "descon", "estadocon", "tipcon", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(txtCodigo, Trim$(txtDescripcion), s_Estado, s_Tipo, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcon")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plconcesunat", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plconcesunat", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fTablaSistema.dcaRegistro, fTablaSistema.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fTablaSistema.dcaRegistro.Recordset.Find ("codcon='" & s_Sunat & "'")
  ' si es actualización pasa al modo visualización
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
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 4060: Me.Width = 7290
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Codigos Sunat"
  lblTitle = "Codigos Sunat"
  n_IndexHelp = -1

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera


  ' Obtengo el modo de operación del registro
  Me.Tag = fTablaSistema.Tag
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For i = 0 To 2
    aElemento(i, 1) = Choose(i + 1, "anadir", "borrar", "modifica")
    aElemento(i, 2) = Choose(i + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
  Next i
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For i = 0 To 3
    aElemento(i, 1) = Choose(i + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(i, 2) = Choose(i + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next i
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fTablaSistema.dcaRegistro.Recordset.EOF Or fTablaSistema.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fTablaSistema.dcaRegistro.Recordset!codcon
  
    
  For i = 0 To 2
    cmbtipo.AddItem Choose(i + 1, "INGRESO", "DESCUENTO", "APORTE")
  Next i
  
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
      MsgBox "Debe Ingresar el Código del " & lblTitle, vbExclamation
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
    cmbtipo.SetFocus
    KeyAscii = 0
  End If
End Sub
