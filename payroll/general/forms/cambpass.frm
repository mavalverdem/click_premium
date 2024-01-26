VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fCambioPassword 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   4350
   Icon            =   "cambpass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   4350
   Begin TabDlg.SSTab tabRegister 
      Height          =   2205
      Left            =   75
      TabIndex        =   9
      Top             =   600
      Width           =   4200
      _ExtentX        =   7408
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
      TabPicture(0)   =   "cambpass.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPassword(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPassword(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPassword(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2205
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1335
         Width           =   1380
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2205
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   975
         Width           =   1380
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2205
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   615
         Width           =   1380
      End
      Begin VB.Label lblDato 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario : "
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
         Height          =   300
         Index           =   0
         Left            =   720
         TabIndex        =   0
         Top             =   255
         Width           =   2865
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirme Password :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   555
         TabIndex        =   5
         Top             =   1380
         Width           =   1545
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Nuevo Password :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   555
         TabIndex        =   3
         Top             =   1020
         Width           =   1545
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Password Anterior :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   555
         TabIndex        =   1
         Top             =   660
         Width           =   1545
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4350
      _Version        =   65536
      _ExtentX        =   7673
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
         Index           =   1
         Left            =   3645
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
         Picture         =   "cambpass.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   3255
         TabIndex        =   7
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "cambpass.frx":0044
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
         TabIndex        =   11
         Top             =   120
         Width           =   2730
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   2895
      Width           =   4350
      _Version        =   65536
      _ExtentX        =   7673
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
Attribute VB_Name = "fCambioPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle, y parametro de codigo
Sub ShowScreen()
  ' Habilita o Inabilita los Controles de Acuerdo a la Acción
  cmdAction(0).Enabled = True: cmdAction(1).Enabled = True
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  lblDato(0) = "Usuario : " & ps_Usuario
  gdl_Procedure.EditText "AT", txtPassword(0), "", "A", False, 10
  gdl_Procedure.EditText "AT", txtPassword(1), "", "A", False, 10
  gdl_Procedure.EditText "AT", txtPassword(2), "", "A", False, 10
End Sub
Private Sub cmdAction_Click(Index As Integer)
Dim s_ClaveUsuario As String                            ' Codigo del registro

  If Index = 0 Then
    ' Realizo las validaciones de los campos a actualizar
    If txtPassword(0) = "" Then Beep: MsgBox "Debe Ingresar Password Anterior", vbExclamation: txtPassword(0).SetFocus: Exit Sub
    If txtPassword(1) = "" Then Beep: MsgBox "Debe Ingresar Nuevo Password", vbExclamation: txtPassword(1).SetFocus: Exit Sub
    If txtPassword(2) = "" Then Beep: MsgBox "Debe Confirmar Nuevo Password", vbExclamation: txtPassword(2).SetFocus: Exit Sub
    If txtPassword(1) <> txtPassword(2) Then Beep: MsgBox "Confirmación de Password debe ser igual al Nuevo Password", vbExclamation: txtPassword(2).SetFocus: Exit Sub
      
    ' Cargo los datos del usuario en un recordset
    s_ClaveUsuario = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, ps_Usuario, "PW")
    
    ' Obtengo los datos del usuario
    If Not (s_ClaveUsuario = "???" Or s_ClaveUsuario = "") Then
      If Trim(s_ClaveUsuario) = txtPassword(0) Then
        Beep
        If MsgBox("¿ Estás Seguro de Cambiar el Password del Usuario '" & ps_Usuario & "' ? ", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
          ' Coloco el puntero en espera
          gdl_Procedure.PunteroEnEspera
          ' Creo los arreglos para la actualización
          a_Campos = Array("codusr", "clausr")
          a_Valores = Array(ps_Usuario, Trim(txtPassword(1)))
          a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter)
          a_Where = Array("codusr")
          
          '[ Inicio la conexión a la base de datos ]
          ps_StrgConnec = OpenConnection(ps_Servidor, ps_BDSystems)
          
          gdl_Conexion.IniciaTransaccion    ' Inicia transacción
          ' Realizo el proceso de actualización del registro
          If Not Records_Upd("sgusr", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
          gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
          MsgBox "Se Actualizo exitosamente el paswword del usuario " & ps_Usuario, vbInformation
          Index = 1
        End If
      Else
        Beep
        MsgBox "Password Anterior Incorrecto", vbCritical
      End If
    Else
      Beep
      MsgBox "Usuario No Registrado ", vbCritical
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
  ' salgo del formulario
  If Index = 1 Then Unload Me

End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 3885: Me.Width = 4440
  gdl_Procedure.CentraFormulario Me
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización de Password"
  lblTitle = "Password"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(2, 1) = "edit": aElemento(2, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 1
      aElemento(n_Index, 1) = Choose(n_Index + 1, "aceptar", "cancelar")
      aElemento(n_Index, 2) = Choose(n_Index + 1, "Actualizar Información de  ", "Cancelar Información de ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(1).Cancel = True
  
  ' Carga los datos en el formulario
  ShowScreen
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub txtPassword_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPassword(Index)
End Sub
Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
