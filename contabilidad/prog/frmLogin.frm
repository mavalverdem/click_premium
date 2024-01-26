VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
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
      TabIndex        =   5
      Tag             =   "Cancelar"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   360
      Left            =   480
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblTexto 
      Caption         =   "&Contraseña:"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "&Contraseña:"
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblTexto 
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
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'Private Declare Function TileWindows Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As Rect, ByVal cKids As Long, lpKids As Long) As Integer
'Private Declare Function AbortPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
'Private Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
'Private Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
'Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
'Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public ubCorrecto As Boolean
Private pocnnBDC As ADODB.Connection
Private pocnnBDS As ADODB.Connection
Private porstAcceso As ADODB.Recordset
Private aIntento()            ' Acceso al sistema

Private Sub Form_Load()
  Dim dsBuffer As String
  Dim dnSize As Long
  ReDim aIntento(2, 0)

  If Not gbEsUsr Then
    txtUserName.Text = "ADMIN"
    txtPassword.Text = "ADMIN"
  Else
    dsBuffer = Space$(255)
    dnSize = Len(dsBuffer)
    GetUserName dsBuffer, dnSize
    If dnSize > 0 Then
      txtUserName.Text = Left$(dsBuffer, dnSize)
    Else
      txtUserName.Text = vbNullString
    End If
  End If
  
  Set pocnnBDC = New Connection
  Set pocnnBDS = New Connection
  Set porstAcceso = New Recordset
  pocnnBDC.CursorLocation = adUseClient
  pocnnBDS.CursorLocation = adUseClient
  With porstAcceso
    .Source = "SELECT SGUsr.CodUsr, SGUsr.AbvUsr, SGUsr.ClaUsr, SGUsr.NomUsr, SGUsr.MesUsr, "
    .Source = .Source & "SGUsr.AnoUsr, SGUsr.EmpUsr, SGUsr.EstUsr, SGUsr.NvlUsr, "
    .Source = .Source & "TGEmp.RazEmp, TGEmp.RUCEmp "
    .Source = .Source & "FROM TGEmp INNER JOIN SGUsr ON TGEmp.CodEmp = SGUsr.EmpUsr "
    .Source = .Source & "WHERE SGUsr.EstUsr IN ('A', 'B') "
    .Source = .Source & "ORDER BY SGUsr.CodUsr"
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
  End With
  
  '[ cambiar las etiquetas
  ReDim aLabel(2, 2)
  Me.Caption = Choose(gsIdioma, "Inicio de sesión", "Start of session")
  For dnSize = 0 To 1
    aLabel(dnSize, 0) = Choose(dnSize + 1, "Usuario", "Contraseña")
    aLabel(dnSize, 1) = Choose(dnSize + 1, "User", "Password")
  Next dnSize
  CaptionBotones Me, True, True, False, False, False, False, False, False, False, False, False, False, False, aLabel

End Sub

Private Sub Form_Activate()

   pocnnBDC.ConnectionString = CONNSTRG & gsNomBDC & ";"
   pocnnBDC.Open
   porstAcceso.ActiveConnection = pocnnBDC
   porstAcceso.Open
   
   If Not gbEsUsr Then
      cmdAceptar_Click
   Else
      txtPassword.SetFocus
   End If
   
End Sub

Private Sub Form_Deactivate()
   porstAcceso.Close
   pocnnBDC.Close
   pocnnBDS.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set porstAcceso = Nothing
   Set pocnnBDC = Nothing
   Set pocnnBDS = Nothing
End Sub

Private Sub cmdAceptar_Click()
  Dim nExpresion As Integer
  
  txtUserName = UCase(txtUserName)
  With porstAcceso
    .MoveFirst
    .Find "CodUsr = '" & txtUserName & "'", , , adBookmarkFirst
    If .EOF Then
      MsgBox Choose(gsIdioma, "No existe Usuario. Vuelva a intentarlo", " User is incorrect. Try again"), , Me.Caption
      txtPassword.SetFocus
      txtPassword.SelStart = 0
      txtPassword.SelLength = Len(txtPassword.Text)
    ElseIf Not !EstUsr = "A" Then
      MsgBox Choose(gsIdioma, "El Usuario no está Activo o se encuentra Bloqueado. Vuelva a intentarlo", "The User is not Active or is Blocked. Try again"), vbExclamation, Me.Caption
      txtPassword.SetFocus
      txtPassword.SelStart = 0
      txtPassword.SelLength = Len(txtPassword.Text)
    Else
'      If UCase(txtPassword) = UCase(gfEnmasc(!ClaUsr)) Then
      If UCase(txtPassword) = UCase(!ClaUsr) Then
        gsNvlUsr = !NvlUsr
        gsMesAct = !MesUsr
        gsAnoAct = !AnoUsr
        gsCodEmp = !EmpUsr
        gsRazEmp = !RazEmp
        gsRUCEmp = !RUCEmp
        gsCodUsr = !CodUsr
        gsAbvUsr = gfEnmasc(!AbvUsr)
        
        ubCorrecto = True
        Me.Hide
      Else
        '[ Bloqueo usuario
        For nExpresion = 0 To UBound(aIntento, 2)
          If aIntento(1, nExpresion) = txtUserName.Text Then Exit For
        Next nExpresion
        If UBound(aIntento, 2) < nExpresion Then
          ReDim Preserve aIntento(2, nExpresion)
          aIntento(1, nExpresion) = txtUserName.Text
          aIntento(2, nExpresion) = 0
        End If
        aIntento(2, nExpresion) = aIntento(2, nExpresion) + 1
        If aIntento(2, nExpresion) = 3 Then
          pocnnBDC.Execute "UPDATE sgusr SET estusr='B' WHERE codusr='" & txtUserName & "'"
          If porstAcceso.State = adStateOpen Then porstAcceso.Close: porstAcceso.Open
        End If
        ']
        MsgBox Choose(gsIdioma, "La contraseña es incorrecta. Vuelva a intentarlo", "The password is incorrect. Try again"), vbCritical, Me.Caption
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
      End If
    End If
  End With

End Sub

Private Sub cmdCancelar_Click()
   ubCorrecto = False
   Me.Hide
End Sub

Private Sub txtPassword_GotFocus()
   txtPassword.SelStart = 0
   txtPassword.SelLength = txtPassword.MaxLength
End Sub

Private Sub txtUserName_GotFocus()
   txtUserName.SelStart = 0
   txtUserName.SelLength = txtUserName.MaxLength
End Sub
