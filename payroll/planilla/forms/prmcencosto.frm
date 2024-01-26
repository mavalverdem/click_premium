VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fPrmCentroCosto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "prmcencosto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6420
   Begin TabDlg.SSTab tabRegister 
      Height          =   2595
      Left            =   75
      TabIndex        =   9
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4577
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   512
      TabMaxWidth     =   3263
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
      TabPicture(0)   =   "prmcencosto.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "shpCuadro(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtLinea"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSegmento"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCliente"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   240
         TabIndex        =   5
         Top             =   1665
         Width           =   980
      End
      Begin VB.TextBox txtSegmento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtLinea 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblDato 
         Caption         =   "Cliente de Negocio :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1425
         Width           =   2715
      End
      Begin VB.Shape shpCuadro 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00C00000&
         FillColor       =   &H00400000&
         Height          =   2055
         Index           =   0
         Left            =   105
         Shape           =   4  'Rounded Rectangle
         Top             =   105
         Width           =   6060
      End
      Begin VB.Label lblDato 
         Caption         =   "Segmento de Negocio :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2715
      End
      Begin VB.Label lblDato 
         Caption         =   "Línea de Negocio :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2715
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   6
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
         TabIndex        =   10
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
         Picture         =   "prmcencosto.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   5400
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
         Picture         =   "prmcencosto.frx":0044
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
         TabIndex        =   7
         Top             =   120
         Width           =   4800
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   3270
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
End
Attribute VB_Name = "fPrmCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String                            ' Codigo del registro
'[
Sub ShowScreen()
    
  ' Información de configuración
  s_Sql = "SELECT lineanegocio, segmentonego, clientenego "
  s_Sql = s_Sql & "FROM plcfgcencosto "
  s_Sql = s_Sql & "WHERE codcco='" & fCentroCosto.dcaRegistro.Recordset!codcco & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    Me.Tag = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtLinea, gdl_Funcion.aTexto(porstRecordset!lineanegocio), Me.Tag, False, porstRecordset!lineanegocio.DefinedSize
    gdl_Procedure.EditText "AT", txtSegmento, gdl_Funcion.aTexto(porstRecordset!segmentonego), Me.Tag, False, porstRecordset!segmentonego.DefinedSize
    gdl_Procedure.EditText "AT", txtCliente, gdl_Funcion.aTexto(porstRecordset!clientenego), Me.Tag, False, porstRecordset!clientenego.DefinedSize
  Else
    Me.Tag = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtLinea, "", Me.Tag, False, porstRecordset!lineanegocio.DefinedSize
    gdl_Procedure.EditText "AT", txtSegmento, "", Me.Tag, False, porstRecordset!segmentonego.DefinedSize
    gdl_Procedure.EditText "AT", txtCliente, "", Me.Tag, False, porstRecordset!clientenego.DefinedSize
  End If

End Sub
']
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  
  ' Realizo las validaciones de los campos a actualizar
  If txtLinea.Text = "" Then Beep: MsgBox "Debe Ingresar Linea de Negocio Centro de Costo", vbExclamation: txtLinea.SetFocus: Exit Sub
  If txtSegmento.Text = "" Then Beep: MsgBox "Debe Ingresar Segmento de Negocio Centro de Costo", vbExclamation: txtSegmento.SetFocus: Exit Sub
  If txtCliente.Text = "" Then Beep: MsgBox "Debe Ingresar Cliente de Negocio Centro Costo", vbExclamation: txtCliente.SetFocus: Exit Sub
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcco", "lineanegocio", "segmentonego", "clientenego", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(fCentroCosto.dcaRegistro.Recordset!codcco, txtLinea.Text, txtSegmento.Text, txtCliente.Text, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcco")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plcfgcencosto", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plcfgcencosto", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' finalizo actualización
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

  'Establece posición y titulo del formulario
  Me.Height = 4200: Me.Width = 6510
  Me.Left = 3580: Me.Top = 2500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Información Negocio Centro Costo"
  lblTitle = "Negocio Centro Costo"
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "edit": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "aceptar"
  aElemento(0, 2) = "Actualizar Información de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdUpdate, aElemento
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Carga los datos en el formulario
  ShowScreen
 
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub txtCliente_GotFocus()
  gdl_Procedure.MarcaGet txtCliente
End Sub
Private Sub txtCliente_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtLinea_GotFocus()
  gdl_Procedure.MarcaGet txtLinea
End Sub
Private Sub txtLinea_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtSegmento_GotFocus()
  gdl_Procedure.MarcaGet txtSegmento
End Sub
Private Sub txtSegmento_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
