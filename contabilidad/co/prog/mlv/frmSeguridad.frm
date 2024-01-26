VERSION 5.00
Begin VB.Form frmSeguridad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguridad de Proceso"
   ClientHeight    =   1545
   ClientLeft      =   4305
   ClientTop       =   6195
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Inicio de sesión"
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   2100
      TabIndex        =   3
      Tag             =   "Cancelar"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   2
      Tag             =   "Aceptar"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   5
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "&Contraseña:"
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Usuari&o:"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Tag             =   "Usuari&o:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pocnnMain As ADODB.Connection
Private porstAcceso As ADODB.Recordset

Private Sub Form_Load()
   
   Set pocnnMain = New ADODB.Connection
   Set porstAcceso = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With porstAcceso
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodUsr, ClaUsr, EstUsr, NvlUsr " _
              & "FROM SGUsr"
      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   
   frmPTraInf.lblEliminar.Tag = ESTCTA_INA
   txtPassword.MaxLength = 10
   txtUserName.Text = gsCodUsr
   txtUserName.Locked = True
   txtUserName.BackColor = &HC0C0C0
 ']
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
  porstAcceso.Close
  pocnnMain.Close
  Set porstAcceso = Nothing
  Set pocnnMain = Nothing
End Sub

Private Sub cmdAceptar_Click()
   
  txtUserName = UCase(txtUserName)
  With porstAcceso
    .MoveFirst
    .Find "CodUsr = '" & txtUserName & "'", , , adBookmarkFirst
    If .EOF Then
      MsgBox "No existe Usuario. Vuelva a intentarlo", , Me.Caption
      txtPassword.SetFocus
    ElseIf Not !NvlUsr = NvlUsr_Adm Then
       MsgBox "El Usuario no puede generar procesos; Vuelva a intentarlo", , Me.Caption
       txtPassword.SetFocus
    ElseIf Not !EstUsr = "A" Then
       MsgBox "El Usuario no está Activo. Vuelva a intentarlo", , Me.Caption
       txtPassword.SetFocus
    Else
      If (UCase(txtPassword) = UCase(!ClaUsr)) Then
        frmPTraInf.lblEliminar.Tag = ESTCTA_ACT
        Unload Me
      Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", vbCritical, Me.Caption
        txtPassword.SetFocus
      End If
    End If
  End With

End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = txtPassword.MaxLength
End Sub

