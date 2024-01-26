VERSION 5.00
Begin VB.Form frmOCla 
   Caption         =   "Cambio de Clave"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      ForeColor       =   &H80000002&
      Height          =   675
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   2760
      Begin VB.TextBox txtClaAct 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1455
         TabIndex        =   0
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Calve Actual :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame fraClaNue 
      ForeColor       =   &H80000002&
      Height          =   675
      Left            =   60
      TabIndex        =   7
      Top             =   900
      Width           =   4755
      Begin VB.TextBox txtClaNueConf 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtClaNue 
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Confirmación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   2
         Left            =   2460
         TabIndex        =   9
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Clave:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   300
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1680
      ScaleHeight     =   690
      ScaleWidth      =   1575
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1575
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   60
         Picture         =   "frmOCla.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   800
         Picture         =   "frmOCla.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOCla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Private uorstMain As ADODB.Recordset

Private Sub Form_Load()
   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodUsr, ClaUsr "
      .Source = .Source & "FROM SGUsr"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   
   With txtClaAct
      .MaxLength = 10
   End With
   With txtClaNue
      .Enabled = False
      .MaxLength = 10
   End With
   With txtClaNueConf
      .Enabled = False
      .MaxLength = 10
   End With
 ']
   cmdGrabar.Enabled = False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Clave Actual :", "Nueva Clave :", "Confirmación :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Actual Password:", "New Password :", "Confirmation :")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, True, False, True, aLabel
 ']
 
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   uorstMain.Close
   uocnnMain.Close
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err
    
   If txtClaNue.Text = txtClaNueConf.Text Then
'      uorstMain!ClaUsr = gfEnmasc(txtClaNue.Text)
      uorstMain!ClaUsr = txtClaNue.Text
      uorstMain.Update
      Unload Me
   End If
   
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub txtClaAct_GotFocus()
   txtClaAct.SelStart = 0
   txtClaAct.SelLength = txtClaNue.MaxLength
End Sub

Private Sub txtClaAct_LostFocus()
   If Len(Trim(txtClaAct.Text)) <> 0 Then
      txtClaAct.Enabled = False
      txtClaNue.Enabled = True
      txtClaNueConf.Enabled = True
      txtClaNue.SetFocus
      txtClaNue.Text = ""
   End If
End Sub

Private Sub txtClaAct_Validate(Cancel As Boolean)
   If Len(Trim(txtClaAct.Text)) = 0 Then
      MsgBox Choose(gsIdioma, "Debe ingresar la clave actual.", "You Should enter actual password"), vbExclamation
      Cancel = True
      Exit Sub
   ElseIf Len(Trim(txtClaAct.Text)) < 5 Then
      MsgBox Choose(gsIdioma, "La clave debe tener, por lo menos, cinco (5) caracteres.", "The password must have at least five (5) characters."), vbExclamation
      Cancel = True
      Exit Sub
   Else 'If Len(Trim(txtClaAct.Text)) <> 0 Then
      With uorstMain
         .MoveFirst
         .Find "CodUsr='" & gsCodUsr & "'"
         If Not .EOF Then
'            If UCase(gfEnmasc(!ClaUsr)) <> UCase(txtClaAct.Text) Then
            If UCase(!ClaUsr) <> UCase(txtClaAct.Text) Then
               MsgBox Choose(gsIdioma, "Clave Inválida", "The password is incorrect"), vbExclamation
               Cancel = True
               Exit Sub
            End If
         Else
            MsgBox Choose(gsIdioma, "Usuario no encontrado.", "User not found"), vbCritical
            End
         End If
      End With
   End If
End Sub

Private Sub txtClaNue_GotFocus()
   txtClaNue.SelStart = 0
   txtClaNue.SelLength = txtClaNue.MaxLength
End Sub

Private Sub txtClaNue_LostFocus()
   If txtClaNue.Text <> "" Then
      txtClaNueConf.Enabled = True
      txtClaNueConf.SetFocus
   Else
      txtClaNueConf.Text = ""
      txtClaNueConf.Enabled = False
   End If
End Sub

Private Sub txtClaNue_Validate(Cancel As Boolean)
   If Len(Trim(txtClaNue.Text)) = 0 Then
      Exit Sub
   ElseIf Len(Trim(txtClaNue.Text)) < 5 Then
      MsgBox Choose(gsIdioma, "La clave debe tener, por lo menos, cinco (5) caracteres.", "The password must have at least five (5) characters."), vbExclamation
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub txtClaNueConf_GotFocus()
   txtClaNueConf.SelStart = 0
   txtClaNueConf.SelLength = txtClaNueConf.MaxLength
End Sub

Private Sub txtClaNueConf_LostFocus()
   If txtClaNue.Text = txtClaNueConf.Text Then
      cmdGrabar.Enabled = True
      cmdGrabar.SetFocus
   Else
      cmdGrabar.Enabled = False
   End If
End Sub

Private Sub txtClaNueConf_Validate(Cancel As Boolean)
   If Len(Trim(txtClaNueConf.Text)) = 0 Then
      Exit Sub
   ElseIf Len(Trim(txtClaNueConf.Text)) < 5 Then
      MsgBox Choose(gsIdioma, "La clave debe tener, por lo menos, cinco (5) caracteres.", "The password must have at least five (5) characters."), vbExclamation
      Cancel = True
      Exit Sub
   ElseIf txtClaNue.Text <> txtClaNueConf.Text Then
      MsgBox Choose(gsIdioma, "Las claves no son iguales.", "The passwords are not same"), vbCritical
      txtClaNueConf.Text = ""
      Cancel = True
      Exit Sub
   End If
End Sub
