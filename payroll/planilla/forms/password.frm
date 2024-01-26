VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form fPassword 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   3480
   ClientTop       =   2805
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "password.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame frmCuadro 
      Height          =   4305
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5070
      _Version        =   65536
      _ExtentX        =   8943
      _ExtentY        =   7594
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Enabled         =   0   'False
      Begin VB.PictureBox pctImagen 
         AutoSize        =   -1  'True
         Height          =   720
         Left            =   3240
         ScaleHeight     =   660
         ScaleWidth      =   750
         TabIndex        =   26
         Top             =   2400
         Visible         =   0   'False
         Width           =   810
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   3495
         TabIndex        =   6
         Top             =   1500
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   2
         Top             =   510
         Width           =   1340
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   840
         Width           =   1340
      End
      Begin VB.PictureBox pctLogo 
         AutoSize        =   -1  'True
         Height          =   3450
         Index           =   0
         Left            =   120
         ScaleHeight     =   3390
         ScaleWidth      =   1725
         TabIndex        =   9
         Top             =   240
         Width           =   1785
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   3735
         TabIndex        =   8
         Top             =   3750
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Cancelar"
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
      Begin Threed.SSCommand cmdOk 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Top             =   3750
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&OK"
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
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2745
         TabIndex        =   1
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2595
         TabIndex        =   3
         Top             =   855
         Width           =   780
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2835
         TabIndex        =   5
         Top             =   1575
         Width           =   540
      End
      Begin VB.Image imgSeguro 
         Height          =   525
         Left            =   1980
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgLogo 
         Height          =   435
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3765
         Width           =   1785
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   4305
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   5070
      _Version        =   65536
      _ExtentX        =   8943
      _ExtentY        =   7594
      _StockProps     =   14
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Enabled         =   0   'False
      Begin VB.TextBox txtLicencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   45
         TabIndex        =   22
         Top             =   2670
         Width           =   1455
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3495
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1335
         Width           =   1455
      End
      Begin VB.TextBox txtDatabase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin VB.PictureBox pctLogo 
         AutoSize        =   -1  'True
         Height          =   3480
         Index           =   1
         Left            =   120
         ScaleHeight     =   3420
         ScaleWidth      =   2160
         TabIndex        =   25
         Top             =   240
         Width           =   2220
      End
      Begin VB.TextBox txtUserid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   18
         Top             =   2010
         Width           =   1455
      End
      Begin VB.TextBox txtClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   2340
         Width           =   1455
      End
      Begin VB.TextBox txtProvider 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3495
         MaxLength       =   25
         TabIndex        =   12
         Top             =   990
         Width           =   1455
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   375
         Index           =   1
         Left            =   3735
         TabIndex        =   24
         Top             =   3750
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Cancelar"
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
      Begin Threed.SSCommand cmdOk 
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   23
         Top             =   3750
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&OK"
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
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licencia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   2685
         TabIndex        =   21
         Top             =   2700
         Width           =   690
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2700
         TabIndex        =   13
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Datos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   2460
         TabIndex        =   15
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   2595
         TabIndex        =   19
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Id :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   2820
         TabIndex        =   17
         Top             =   2040
         Width           =   555
      End
      Begin VB.Image imgLogo 
         Height          =   435
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3765
         Width           =   1785
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   2550
         TabIndex        =   11
         Top             =   1035
         Width           =   825
      End
   End
End
Attribute VB_Name = "fPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                     ' Declarar variable antes de usarla

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private s_Archivo As String, s_ToolText As String   ' Archivo y mensaje de imagen, icono o control al posicionar mouse
Private i As Byte                                   ' Indice para bucle
Private Sub cmdExit_Click(Index As Integer)
  Unload Me
End Sub
Private Sub cmdOk_Click(Index As Integer)
    
Dim pnSize As Long
Dim nFile As Integer, s_Linea As String
  
  If n_SwConfigura = 0 Then
    If txtUser = "" Then Beep: MsgBox "Debe Ingresar Usuario", vbExclamation: txtUser.SetFocus: Exit Sub
    If txtPassword = "" Then Beep: MsgBox "Debe Ingresar Password", vbExclamation: txtPassword.SetFocus: Exit Sub
    If Not gdl_Funcion.ValidaFecha(Format$(mskFecha, s_FormatoFecha), 2000) Then mskFecha.SetFocus: Exit Sub
    
    '[ Conexón a la Base de Datos y Servidor
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_BDSystems)
    ']
       
    ' Cargo los datos del usuario en un recordset
    s_Sql = "SELECT codusr, clausr, empusr, anousr, mesusr, nvlusr "
    s_Sql = s_Sql & " FROM sgusr "
    s_Sql = s_Sql & " WHERE codusr='" & Trim$(txtUser) & "'"
    s_Sql = s_Sql & " AND estusr='A' "
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    
    ' Obtengo los datos del usuario
    
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      If UCase(porstRecordset!clausr) = UCase(txtPassword) Then
        imgSeguro.Picture = LoadPicture()
        s_Archivo = gdl_Procedure.ps_PathImagen & "seguroac.ico"
        If gdl_Funcion.ExisteArchivo(s_Archivo) Then
          imgSeguro.Picture = LoadPicture(s_Archivo)
        End If
        imgSeguro.Refresh
        Beep
        ' Capturo Nombre de Usuario y Tipo de Cambio
        ps_Usuario = Trim$(txtUser)
        ps_Anyo = porstRecordset!anousr
        ps_Mes = porstRecordset!mesusr
        ps_NivelUsr = porstRecordset!nvlusr
        pl_Salir = True
        Unload Me
      Else
        Beep
        MsgBox "Password Incorrecto", vbCritical
        txtPassword = "": txtPassword.SetFocus
      End If
    Else
      Beep
      MsgBox "Usuario No Registrado " & Err.Description, vbCritical
      txtUser = "": txtPassword = "": txtUser.SetFocus
    End If
    ' Cierro Tabla de Usuario
    Set porstRecordset = Nothing
  Else
    If txtProvider = "" Then Beep: MsgBox "Debe Ingresar Nombre del Proveedor de Servicio", vbExclamation: txtProvider.SetFocus: Exit Sub
    If txtServer = "" Then Beep: MsgBox "Debe Ingresar Nombre del Servidor", vbExclamation: txtServer.SetFocus: Exit Sub
    If txtDatabase = "" Then Beep: MsgBox "Debe Ingresar Nombre de la Base de Datos", vbExclamation: txtDatabase.SetFocus: Exit Sub
    If txtUserid = "" Then Beep: MsgBox "Debe Ingresar Nombre del Usuario", vbExclamation: txtUserid.SetFocus: Exit Sub
    If txtLicencia = "" Then Beep: MsgBox "Debe Ingresar Nombre de la Empresa a Licenciar", vbExclamation: txtLicencia.SetFocus: Exit Sub
    ' Abro Archivo de Texto
    nFile = FreeFile
    s_Archivo = ps_WinSystem & "\" & pFileSystem
      
    Open s_Archivo For Output Access Write Lock Read Write As #nFile
    For i = 0 To 7
      ' Diseño la Linea a Grabar
      s_Linea = ""
      s_Linea = s_Linea & Choose(i + 1, "{Configuracion}", "[Planilla]", "[Proveedor]=", "[Servidor]=", "[UserId]=", "[Password]=", "[BaseDatos]=", "[Licencia]=")
      s_Linea = s_Linea & gdl_Funcion.Encripta(gdl_Funcion.aTexto(Choose(i + 1, "", "", txtProvider, txtServer, txtUserid, txtClave, txtDatabase, txtLicencia)))
      ' Grabo la Linea en el Archivo
      Print #nFile, s_Linea
    Next i
    ' Cierro Archivo de Texto
    Close #nFile
    Unload Me
  End If

End Sub
Private Sub Form_Load()
    
  Dim psWinUser As String, pnSize As Long

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  pl_Salir = False
  
  ' Verifico que exista el icono del formulario
  Me.Icon = LoadPicture()
  s_Archivo = gdl_Procedure.ps_PathImagen & IIf(n_SwConfigura = 0, "seguro.ico", "configura.ico")
  If gdl_Funcion.ExisteArchivo(s_Archivo) Then
    Me.Icon = LoadPicture(s_Archivo)
  End If
  ' Actualizo el Titulo del Formulario
  Me.Caption = Choose(n_SwConfigura + 1, "Seguridad del Sistema", "Configuración - Parametros de Conexión")
  ' Visualiso el Cuadro y su Ubicación
  frmCuadro(n_SwConfigura).Visible = True
  frmCuadro(n_SwConfigura).Top = 0
  frmCuadro(n_SwConfigura).Enabled = True
  ' Verifico que exista el Logo de Ingreso
  s_Archivo = gdl_Procedure.ps_PathImagen & Choose(n_SwConfigura + 1, "logo ingreso", "configura") & ".jpg"
  If dir$(s_Archivo, vbNormal) <> "" Then
      pctLogo(n_SwConfigura).Picture = LoadPicture(s_Archivo)
  End If
  pctLogo(n_SwConfigura).Refresh
  
  ' Verifico que exista el logo de sistemas
  s_Archivo = gdl_Procedure.ps_PathImagen & "logo sysma.jpg"
  If dir$(s_Archivo, vbNormal) <> "" Then
      imgLogo(n_SwConfigura).Picture = LoadPicture(s_Archivo)
  End If
  imgLogo(n_SwConfigura).Refresh
  
  If n_SwConfigura = 0 Then
    ' Verifico que exista el Icono de Seguridad
    imgSeguro.Picture = LoadPicture()
    s_Archivo = gdl_Procedure.ps_PathImagen & "seguroin.ico"
    If dir$(s_Archivo, vbNormal) <> "" Then
        imgSeguro.Picture = LoadPicture(s_Archivo)
    End If
    imgSeguro.Refresh
    
    pctImagen.Picture = LoadPicture()
    s_Archivo = gdl_Procedure.ps_PathImagen & "animacion.gif"
    If dir$(s_Archivo, vbNormal) <> "" Then
        pctImagen.Picture = LoadPicture(s_Archivo)
    End If
    pctImagen.Refresh
    
    ' Recupero el usuario de windows
    psWinUser = Space$(255)
    pnSize = Len(psWinUser)
    GetUserName psWinUser, pnSize
    txtUser.Text = IIf(pnSize > 0, Left$(psWinUser, pnSize), vbNullString)
    
    ' Inicializo fecha
    mskFecha.Mask = "": mskFecha = Format(gs_FechaHora, s_FormatoFecha): mskFecha.Mask = "##/##/####"
  End If
  '  Inicializo la Posicion del Formulario
  gdl_Procedure.CentraFormulario Me
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set gdl_Conexion = Nothing
  imgSeguro.Picture = LoadPicture()
End Sub
Private Sub mskFecha_GotFocus()
  gdl_Procedure.MarcaGet mskFecha
End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And mskFecha.ClipText <> "" Then cmdOk(0).SetFocus
End Sub
Private Sub txtClave_GotFocus()
  gdl_Procedure.MarcaGet txtClave
End Sub
Private Sub txtClave_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    txtLicencia.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtDatabase_GotFocus()
  gdl_Procedure.MarcaGet txtDatabase
End Sub
Private Sub txtDatabase_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn And Len(txtDatabase) > 0 Then
    txtUserid.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtLicencia_GotFocus()
  gdl_Procedure.MarcaGet txtLicencia
End Sub
Private Sub txtLicencia_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn And Len(txtPassword) > 0 Then
    cmdOk(1).SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtPassword_GotFocus()
  gdl_Procedure.MarcaGet txtPassword
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn And Len(txtPassword) > 0 Then
    cmdOk(0).SetFocus
    KeyAscii = 0
  End If
    
End Sub
Private Sub txtProvider_GotFocus()
  gdl_Procedure.MarcaGet txtProvider
End Sub
Private Sub txtProvider_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn And Len(txtProvider) > 0 Then
    txtServer.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtServer_GotFocus()
  gdl_Procedure.MarcaGet txtServer
End Sub
Private Sub txtServer_KeyPress(KeyAscii As Integer)
    
  If KeyAscii = vbKeyReturn And Len(txtServer) > 0 Then
    txtDatabase.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtUser_GotFocus()
  gdl_Procedure.MarcaGet txtUser
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn And Len(txtUser) > 0 Then
    txtPassword.SetFocus
    KeyAscii = 0
  End If
    
End Sub
Private Sub txtUserid_GotFocus()
  gdl_Procedure.MarcaGet txtUserid
End Sub
Private Sub txtUserid_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn And Len(txtUserid) > 0 Then
    txtClave.SetFocus
    KeyAscii = 0
  End If

End Sub
