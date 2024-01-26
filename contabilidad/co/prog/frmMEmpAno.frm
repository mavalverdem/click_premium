VERSION 5.00
Begin VB.Form frmMEmpAno 
   Caption         =   "[Entidad]"
   ClientHeight    =   2355
   ClientLeft      =   165
   ClientTop       =   1350
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame fraCuadro 
      Height          =   1320
      Left            =   915
      TabIndex        =   0
      Top             =   960
      Width           =   3270
      Begin VB.ComboBox cmbEjercicio 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   285
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   495
         Width           =   2610
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   1
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   5160
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
         Left            =   4350
         Picture         =   "frmMEmpAno.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
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
         Left            =   0
         Picture         =   "frmMEmpAno.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   375
      TabIndex        =   6
      Top             =   675
      Width           =   780
   End
End
Attribute VB_Name = "frmMEmpAno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nContador As Integer

Private Function GeneraEjercicio(ByVal sEmpresa As String, ByVal sEjercicio As String, ByVal sArchivo As String) As Boolean
  Dim pofsoFileQry As FileSystemObject, potxtFileQry As TextStream
  Dim s_Catalogo As String, sTabla As String, sPeriodo As String
  Dim nGenera As Integer, nSecuencia As Integer
  Dim psLinea As String, psSentencia As String
  Dim nRegistro As Long, nRegistros As Long
  Dim cnnNuevoAno As New Connection
  
  On Error GoTo Errores
  
  s_Catalogo = sEjercicio
  
  '[ Inicio la conexión a la base de datos ]
  With cnnNuevoAno
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  
  ' Verifico si existe ejercicio
  psSentencia = "SELECT COUNT(a.codemp) AS nExiste "
  psSentencia = psSentencia & "FROM cocfg a, tgcfg b "
  psSentencia = psSentencia & "WHERE a.codemp='" & sEmpresa & "' "
  psSentencia = psSentencia & "AND a.pdoano='" & s_Catalogo & "' "
  psSentencia = psSentencia & "AND b.codemp=a.codemp "
  psSentencia = psSentencia & "AND b.pdoano=a.pdoano"
  nRegistros = CInt(gfRetornaValor(CONNSTRG & gsNomBDS, psSentencia))
  If nRegistros = 1 Then MsgBox Choose(gsIdioma, "Existe año ", "Exist the year ") & Right(cmbEjercicio.Text, 4) & Choose(gsIdioma, " de la empresa ", " of the company ") & frmMEmpGrd.uorstMain!codemp & ".", vbExclamation: GoTo Finalizar

  ' Verifico si existe empresa
  nGenera = 1
  psSentencia = "SELECT COUNT(a.codemp) AS nExiste "
  psSentencia = psSentencia & "FROM cocfg a, tgcfg b "
  psSentencia = psSentencia & "WHERE a.codemp='" & sEmpresa & "' "
  psSentencia = psSentencia & "AND b.codemp=a.codemp"
  nRegistros = CInt(gfRetornaValor(CONNSTRG & gsNomBDS, psSentencia))
  If nRegistros = 0 Then nGenera = 2
  
  ' Elimino y genero las tablas temporales
  nSecuencia = Choose(nGenera, 3, 11)
  For nRegistro = 1 To nSecuencia
    sTabla = Choose(nRegistro, "cocfg", "cociemes", "tgcfg", "cocco", "cocta", "codro", "coefe", "cofjo", "tgtdc", "coefi", "coefilin")
    psSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & Len(sTabla) + 5 & ")='#tmp" & sTabla & "_') DROP TABLE #tmp" & sTabla
    cnnNuevoAno.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp" & sTabla, psSentencia)
    
    psSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmp" & sTabla & " ", "")
    psSentencia = psSentencia & "SELECT * "
    psSentencia = psSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmp" & sTabla & " ", "")
    psSentencia = psSentencia & "FROM " & sTabla & " "
    psSentencia = psSentencia & "WHERE codemp='mav'"
    cnnNuevoAno.Execute psSentencia, nRegistros
  Next nRegistro

  ' Creo objeto de archivo
  Set pofsoFileQry = CreateObject("Scripting.FileSystemObject")
  psSentencia = ""
  cnnNuevoAno.BeginTrans           'Inicia transacción
  sArchivo = ps_WinSystem & "\" & sArchivo
  ' Aperturo el archivo
  Set potxtFileQry = pofsoFileQry.OpenTextFile(sArchivo, ForReading, False, TristateFalse)
  Do While Not potxtFileQry.AtEndOfStream
    psLinea = potxtFileQry.ReadLine
    If Trim(psLinea) = "[Parametros]" Then psLinea = "#"
    ' Verifico si inicializo la información adicional
    If Trim(psLinea) = "[Informacion]" Then
      If nGenera <> 2 Then Exit Do
      psLinea = "#"
    End If
    If Left(psLinea, 1) <> "#" Then
      psSentencia = psSentencia & psLinea
      If Right(Trim(psLinea), 1) = ";" Then
        psSentencia = Replace(psSentencia, "INSERT INTO ", "INSERT INTO " & ps_Prefijo & "tmp")
        cnnNuevoAno.Execute psSentencia, nRegistro
        psSentencia = ""
      End If
    End If
  Loop
  potxtFileQry.Close
  
  ' actualizo informacion temporal
  For nRegistro = 1 To nSecuencia
    ' actualizo empresa y ejercicio
    sTabla = Choose(nRegistro, "cocfg", "cociemes", "tgcfg", "cocco", "cocta", "codro", "coefe", "cofjo", "tgtdc", "coefi", "coefilin")
    sPeriodo = Choose(nRegistro, "S", "S", "S", "S", "S", "S", "S", "S", "N", "S", "S")
    psSentencia = "UPDATE " & ps_Prefijo & "tmp" & sTabla & " "
    psSentencia = psSentencia & "SET codemp='" & sEmpresa & "' "
    psSentencia = psSentencia & IIf(sPeriodo = "S", ", pdoano='" & s_Catalogo & "'", "")
    cnnNuevoAno.Execute psSentencia, nRegistros
    ' Inserto la informacion
    psSentencia = "INSERT INTO " & sTabla & " "
    psSentencia = psSentencia & "SELECT * "
    psSentencia = psSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
    psSentencia = psSentencia & "WHERE codemp='" & sEmpresa & "'"
    cnnNuevoAno.Execute psSentencia, nRegistros
  Next nRegistro
  cnnNuevoAno.CommitTrans           'Confirma transacción
  GeneraEjercicio = True
  GoTo Finalizar
  
Errores:
  cnnNuevoAno.RollbackTrans
  gpErrores
Finalizar:
  ' Reinicializo los mensajes
  Set pofsoFileQry = Nothing
  Set potxtFileQry = Nothing
  '[ Finalizo la conexión a la base de datos ]
  cnnNuevoAno.Close
  Set cnnNuevoAno = Nothing

End Function

Private Sub Form_Load()
  
  '[ Cargo mensajes de botones y etiquetas
  lblEmpresa.Caption = "Empresa : " & Trim(frmMEmpGrd.uorstMain!codemp) & " - " & frmMEmpGrd.uorstMain!RazEmp
  Me.Caption = Choose(gsIdioma, "Ejercicio Contable", "Accounting Period")
  lblTexto(0).Caption = Choose(gsIdioma, "Ejercicio", "Fiscal Year")
  cmdAceptar.Caption = Choose(gsIdioma, "&Aceptar", "&Accept")
  cmdSalir.Caption = Choose(gsIdioma, "&Salir", "&Exit")
  ']
  For nContador = (Val(gsAnoAct) - 5) To (Val(gsAnoAct) + 5)
    cmbEjercicio.AddItem lblTexto(0).Caption & " " & Format(nContador, "0000")
  Next nContador
  cmbEjercicio.ListIndex = 6
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
'   gpTUg_Resize Me
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   cmdSalir.Left = Width - 820
End Sub

Private Sub cmdAceptar_Click()
  Dim sConfiguracion As String, sArchivo As String
    
  On Error GoTo Err
  
  'Mensaje de verificación            'Cambiar.
  If MsgBox(Choose(gsIdioma, "Confirme la creación del año ", "You Confirm the creation of year ") & Right(cmbEjercicio.Text, 4) & Choose(gsIdioma, " de la empresa ", " of the company ") & frmMEmpGrd.uorstMain!codemp & ".", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    
    ' Coloco el puntero en espera
    Screen.MousePointer = vbHourglass
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    ' Verificio la existencia de estructuras
    sArchivo = "cobdqryd" & IIf(ps_Plataforma = pSrvMySql, "m", "s") & ".mai"
    sConfiguracion = ps_WinSystem & "\" & sArchivo
    If StrConv(Dir$(sConfiguracion, vbHidden), vbLowerCase) <> LCase(sArchivo) Then
      MsgBox Choose(gsIdioma, "No existe el archivo de configuración.", "The configuration file does not exist."), vbCritical
      GoTo Err
    End If
    ' Genero la creacion de la nueva empresa
    If Not GeneraEjercicio(frmMEmpGrd.uorstMain!codemp, Right(Trim(cmbEjercicio.Text), 4), sArchivo) Then GoTo Err
    MsgBox Choose(gsIdioma, "Finalizo creación del año ", "finished creation of year ") & Right(cmbEjercicio.Text, 4) & Choose(gsIdioma, " de la empresa ", " of the company ") & frmMEmpGrd.uorstMain!codemp & ".", vbInformation
    
  End If

Err:
  If Err.Number <> 0 Then gpErrores
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  ' Coloco el puntero en normal
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub
