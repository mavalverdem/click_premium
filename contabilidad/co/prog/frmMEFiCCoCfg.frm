VERSION 5.00
Begin VB.Form frmMEFiCCoCfg 
   Caption         =   "[Entidad]"
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   7320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1800
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5910
      Width           =   3480
      Begin VB.CommandButton cmdAvanzar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         Picture         =   "frmMEFiCCoCfg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmdRetroceder 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         Picture         =   "frmMEFiCCoCfg.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   75
         Visible         =   0   'False
         Width           =   360
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
         Left            =   2690
         Picture         =   "frmMEFiCCoCfg.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer"
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
         Left            =   1950
         Picture         =   "frmMEFiCCoCfg.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   720
      End
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
         Left            =   1220
         Picture         =   "frmMEFiCCoCfg.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdCorregir 
         Caption         =   "&Corregir"
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
         Left            =   480
         Picture         =   "frmMEFiCCoCfg.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.TextBox txtDato 
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
      Index           =   0
      Left            =   1290
      TabIndex        =   3
      Top             =   645
      Width           =   5895
   End
   Begin VB.TextBox txtLlave 
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
      Index           =   0
      Left            =   1290
      TabIndex        =   1
      Top             =   105
      Width           =   560
   End
   Begin VB.Frame fraNivel 
      Caption         =   "Nivel de Cuentas"
      ForeColor       =   &H80000002&
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7110
      Begin VB.OptionButton optNivel 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   2
         Left            =   2040
         TabIndex        =   8
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   3
         Left            =   3000
         TabIndex        =   9
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "6 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   4
         Left            =   3960
         TabIndex        =   10
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "7 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   5
         Left            =   4920
         TabIndex        =   11
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "8 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   6
         Left            =   5880
         TabIndex        =   12
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivel 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.TextBox txtNombre 
      Height          =   330
      Left            =   4185
      MaxLength       =   15
      TabIndex        =   18
      Top             =   5505
      Width           =   2070
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Agregar"
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
      Left            =   3330
      Picture         =   "frmMEFiCCoCfg.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   720
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Quitar"
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
      Left            =   3330
      Picture         =   "frmMEFiCCoCfg.frx":0C2E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3525
      Width           =   720
   End
   Begin VB.ListBox lstCOCCoCfg 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   2985
      Left            =   4185
      TabIndex        =   16
      Top             =   2040
      Width           =   3000
   End
   Begin VB.ListBox lstCOCCo 
      BackColor       =   &H00E0E0E0&
      Height          =   3765
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Width           =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   135
      X2              =   7200
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
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
      Left            =   210
      TabIndex        =   2
      Top             =   705
      Width           =   990
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Formato :"
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
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Titulo Columna :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   4185
      TabIndex        =   17
      Top             =   5235
      Width           =   1140
   End
End
Attribute VB_Name = "frmMEFiCCoCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

Private psConnStrgSele As String
Private pnColumnaOrd As Integer

Private psStrgSelect As String, psStrgWhere As String
Private sFormato As String, codigo As String
Private nIndex As Byte, fila As Byte, nNivel As Integer

Private Sub cmdAdd_Click()
  Dim nListIndex As Integer, nListCount As Integer
  
  If lstCOCCoCfg.ListCount >= 11 Then MsgBox "Máximo de columnas posibles", vbExclamation: Exit Sub
  If lstCOCCo.ListCount >= 1 Then
    nListIndex = lstCOCCo.ListIndex
    If nListIndex = -1 Then Beep: MsgBox "Debe seleccionar registro de columna", vbExclamation: lstCOCCo.SetFocus: Exit Sub
    nListCount = lstCOCCo.ListCount - 2
    lstCOCCoCfg.AddItem lstCOCCo.Text
    lstCOCCo.RemoveItem nListIndex
    nListIndex = IIf(nListCount >= nListIndex, nListIndex, nListCount)
    lstCOCCo.ListIndex = nListIndex
  End If
  
End Sub

Public Sub cmdCorregir_Click()
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Private Sub cmdDel_Click()
  Dim nListIndex As Integer, nListCount As Integer
    
  If lstCOCCoCfg.ListCount >= 1 Then
    nListIndex = lstCOCCoCfg.ListIndex
    If nListIndex = -1 Then Beep: MsgBox "Debe seleccionar registro a desagrupar", vbExclamation: lstCOCCoCfg.SetFocus: Exit Sub
    nListCount = lstCOCCoCfg.ListCount - 2
    lstCOCCo.AddItem lstCOCCoCfg.Text
    lstCOCCoCfg.RemoveItem nListIndex
    nListIndex = IIf(nListCount >= nListIndex, nListIndex, nListCount)
    lstCOCCoCfg.ListIndex = nListIndex
  End If

End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdGrabar_Click()
  On Error GoTo Err
  
  If txtDato(0).Text = "" Then Beep: MsgBox Choose(gsIdioma, "Debe Ingresar Descripción del Formato", "You must enter Description the format"), vbExclamation: txtDato(0).SetFocus: Exit Sub
  With frmMEFiCCoGrd                     'Cambiar Formulario de Grid.
    .uocnnMain.BeginTrans            'INICIA TRANSACCION.
    upDatosDesconectados 0
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    ' Actualizo informacion
    .uorstMain_0.Requery
    .uorstMain_1.Requery
    .uorstCCoCfg.Requery
    .upDatosGrid 0
    .upDatosGrid 1
    '[Búsqueda de llave actual.     'Cambiar.
    .uorstMain_0.Find "codcfg='" & txtLlave(0).Text & "'"
    ']
    If pbNuevo Then
      cmdGrabar.Enabled = False
      upHabilitacion False
      
      upDatosPredeterminados
      '[Llave con el foco al añadir.  'Cambiar.
      txtLlave(0).SetFocus
      ']
    Else
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      upHabilitacion False
    End If
  End With
  
  Exit Sub
Err:
  gpErrores
  frmMEFiCCoGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
  
End Sub

'[Propio del formulario.
'Public uorstTGDtt As ADODB.Recordset
']

Private Sub Form_Load()
  Dim sNivel As String
  pbValidada = False
  Me.KeyPreview = True
   
  With frmMEFiCCoGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
    txtLlave(0).MaxLength = .uorstCCoCfg!codcfg.DefinedSize
    ']
    txtDato(0).MaxLength = .uorstCCoCfg!detcfg.DefinedSize
  End With
   
 '[Recordsets                          'Cambiar.
  sFormato = IIf(frmREFiCCo.optFormato(0).Value, 0, 1)
  If sFormato = "0" Then
    psStrgSelect = "SELECT tbl.CodCCo, " & Choose(gsIdioma, "tbl.DetCCo", "tbl.DetCCox") & " AS DetCCo "
    psStrgSelect = psStrgSelect & "FROM cocco tbl "
  Else
    psStrgSelect = "SELECT tbl.CodCta AS CodCCo, " & Choose(gsIdioma, "tbl.DetCta", "tbl.DetCtax") & " AS DetCCo "
    psStrgSelect = psStrgSelect & "FROM cocta tbl "
  End If
  psStrgSelect = psStrgSelect & "WHERE tbl.codemp='" & gsCodEmp & "' "
  psStrgSelect = psStrgSelect & "AND tbl.pdoano='" & gsAnoAct & "' "
 ']
   
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  upHabilitacion False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Formato :", "Descripción :", "Título Columna :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Format :", "Description :", "Title Column :")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, True, False, True, aLabel
  ']
  
  '[ Cargo los mensajes de botones
  fraNivel.Caption = Choose(sFormato + 1, Choose(gsIdioma, "Nivel Centro Costos", "Cost Center Level"), Choose(gsIdioma, "Nivel de Cuentas", "Account Level"))
  optNivel(7).Caption = Choose(gsIdioma, "Detalle", "Detail")
  optNivel(0).Caption = Choose(gsIdioma, "2 dígitos", "2 digits")
  optNivel(1).Caption = Choose(gsIdioma, "3 dígitos", "3 digits")
  optNivel(2).Caption = Choose(gsIdioma, "4 dígitos", "4 digits")
  optNivel(3).Caption = Choose(gsIdioma, "5 dígitos", "5 digits")
  optNivel(4).Caption = Choose(gsIdioma, "6 dígitos", "6 digits")
  optNivel(5).Caption = Choose(gsIdioma, "7 dígitos", "7 digits")
  optNivel(6).Caption = Choose(gsIdioma, "8 dígitos", "8 digits")
  ']
  ' Inicializo los niveles
  sNivel = Choose(sFormato + 1, gsNivCCo, gsNivCta)
  For nIndex = 1 To Len(sNivel)
    optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Visible = True
    Select Case nIndex
     Case Is = 1
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 120
     Case Is = 2
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 1080
     Case Is = 3
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 2040
     Case Is = 4
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 3000
     Case Is = 5
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 3960
     Case Is = 6
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 4920
     Case Is = 7
      optNivel(Val(Mid(sNivel, nIndex, 1)) - 2).Left = 5880
    End Select
  Next
  fraNivel.Width = optNivel(Val(Mid(sNivel, nIndex - 1, 1)) - 2).Left + 1035
   
End Sub

Private Sub Form_Activate()

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
'   Set uocnnMain = Nothing
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

'[Código propio del formulario.

']
Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
   
   'Orden: Corregir.
   zaOpciones = Array(gbPms02)
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdCorregir.Enabled = IIf(pbNuevo, False, taOpciones(0))
End Property
Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  Dim flag As Byte, pnNivRegistro As Byte
  
  If nNivel = 9 Then nNivel = Val(Right(Choose(Val(sFormato) + 1, gsNivCCo, gsNivCta), 1))
  On Error GoTo Err
  With frmMEFiCCoGrd
    If tnFase = 0 Then
      ' Elimino los registros de titulos
      .uocnnMain.Execute "DELETE FROM coccocfg WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND codcfg='" & txtLlave(0).Text & "' AND tipofmt='" & sFormato & "'"
      If .uorstCCoCfg.State = adStateOpen Then .uorstCCoCfg.Close
      .uorstCCoCfg.Open
      flag = 0
      ' Inserto los registros
      If lstCOCCoCfg.ListCount <> 0 Then
        For nIndex = 0 To lstCOCCoCfg.ListCount - 1
          lstCOCCoCfg.ListIndex = nIndex
          With .uorstCCoCfg
            .AddNew
            .Fields!codemp = gsCodEmp
            .Fields!pdoano = gsAnoAct
            .Fields!tipofmt = sFormato
            .Fields!codcfg = txtLlave(0).Text
            .Fields!numord = Format(nIndex, "00")
            .Fields!codcco = Left(lstCOCCoCfg.Text, InStr(lstCOCCoCfg.Text, "-") - 2)
            .Fields!DetCCo = Mid(lstCOCCoCfg.Text, InStr(lstCOCCoCfg.Text, "-") + 2, 15)
            .Fields!nivel = nNivel
            .Fields!detcfg = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
            .Fields!UsrCre = gsAbvUsr
            .Fields!FyHCre = Now
            .Update
            flag = 1
          End With
        Next nIndex
      End If
      
      '[ Graba Registro de Otros ]
      If flag = 1 Then
        With .uorstCCoCfg
          .AddNew
          .Fields!codemp = gsCodEmp
          .Fields!pdoano = gsAnoAct
          .Fields!tipofmt = sFormato
          .Fields!codcfg = txtLlave(0).Text
          .Fields!numord = "XX"
          .Fields!codcco = String(nNivel, "X")
          .Fields!DetCCo = Choose(gsIdioma, "Otros", "Others")
          .Fields!nivel = nNivel
          .Fields!detcfg = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
          .Fields!UsrCre = gsAbvUsr
          .Fields!FyHCre = Now
          .Update
        End With
      End If
    Else
      'Llaves.
      txtLlave(0).Text = .uorstCCoCfg!codcfg
      'Datos.
      txtDato(0).Text = IIf(IsNull(.uorstCCoCfg!detcfg), "", .uorstCCoCfg!detcfg)
      ' Obtengo el nivel definido
      nIndex = 9
      nIndex = CInt(.uorstCCoCfg!nivel)
      optNivel(nIndex - 2).Value = True
      Carga_Listas
    End If
  End With
  Exit Sub
Err:
  gpErrores
  
  Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
  Dim dnContador As Integer

  'Llaves.
  txtLlave(0).Text = ""

  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Text = ""
    Next
  End With
  'Ayudas.
  txtNombre.Text = ""
  optNivel(7).Value = True
  lstCOCCoCfg.Clear

'   lblDatoDeta(2).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Integer

  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  txtNombre.Enabled = tbHabilitar
  fraNivel.Enabled = tbHabilitar
  cmdAdd.Enabled = tbHabilitar
  cmdDel.Enabled = tbHabilitar
End Sub


Private Sub Carga_Listas()
  Dim porsClone As ADODB.Recordset
  
  Set porsClone = New ADODB.Recordset
  Set porsClone = frmMEFiCCoGrd.uorstCCoCfg.Clone()
  
  ' Inicializo lista de configuración
  lstCOCCoCfg.Clear
  If Not (porsClone.EOF And porsClone.BOF) Or porsClone.RecordCount > 0 Then
    porsClone.MoveFirst
    While Not porsClone.EOF
      If Left(porsClone!codcco, 2) <> "XX" Then
        lstCOCCoCfg.AddItem porsClone!codcco & " - " & porsClone!DetCCo
      End If
      porsClone.MoveNext
    Wend
  End If
  porsClone.Close
  Set porsClone = Nothing
       
End Sub

Private Sub lstCOCCoCfg_Click()

  fila = lstCOCCoCfg.ListIndex
  codigo = Left(lstCOCCoCfg.Text, InStr(lstCOCCoCfg.Text, "-") - 2)
  txtNombre.Text = Mid(lstCOCCoCfg.Text, InStr(lstCOCCoCfg.Text, "-") + 2)

End Sub

Private Sub optNivel_Click(Index As Integer)
  
  ' Defino el nivel de  registros
  nNivel = Index + 2
  psStrgWhere = "AND "
  If nNivel = 9 Then
    psStrgWhere = psStrgWhere & Choose(sFormato + 1, IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(tbl.CodCCo))=5 ", "tbl.TpoCta='" & TPOCTA_TRA & "' ")
  Else
    psStrgWhere = psStrgWhere & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(tbl." & Choose(sFormato + 1, "CodCCo", "CodCta") & "))=" & nNivel & " "
  End If
  psStrgWhere = psStrgWhere & "AND NOT EXISTS(SELECT * FROM CoCCoCfg cfg "
  psStrgWhere = psStrgWhere & "WHERE cfg.codemp=tbl.codemp AND cfg.pdoano=tbl.pdoano "
  psStrgWhere = psStrgWhere & "AND cfg.CodCCo=tbl." & Choose(sFormato + 1, "CodCCo", "CodCta") & " "
  psStrgWhere = psStrgWhere & "AND cfg.TipoFmt='" & sFormato & "' "
  psStrgWhere = psStrgWhere & "AND cfg.codcfg='" & txtLlave(0).Text & "') "
  With frmMEFiCCoGrd.uorstCOCta
    If .State = adStateOpen Then .Close
    .Source = psStrgSelect & psStrgWhere
    .Open
  End With
  ' Inicializo lista ayuda
  lstCOCCo.Clear
  If Not (frmMEFiCCoGrd.uorstCOCta.EOF And frmMEFiCCoGrd.uorstCOCta.BOF) Or frmMEFiCCoGrd.uorstCOCta.RecordCount > 0 Then
    While Not frmMEFiCCoGrd.uorstCOCta.EOF
      lstCOCCo.AddItem frmMEFiCCoGrd.uorstCOCta!codcco & " - " & frmMEFiCCoGrd.uorstCOCta!DetCCo
      frmMEFiCCoGrd.uorstCOCta.MoveNext
    Wend
  End If
          
End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
  txtLlave(Index).SelStart = 0
  txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
  If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err

  Dim dvRegistro As Variant
   
  'Valida la llave.                    'Cambiar.
  If Len(Trim(txtLlave(Index).Text)) <> 0 Then
    With frmMEFiCCoGrd.uorstMain_0
      If Not (.BOF And .EOF) Then
        dvRegistro = .Bookmark
        .MoveFirst
        .Find "codcfg='" & txtLlave(0).Text & "'"
        If Not .EOF Then
          MsgBox TEXT_8007, vbExclamation
          If dvRegistro <> -1 Then .Bookmark = dvRegistro
          Cancel = True
          Exit Sub
        End If
        .Bookmark = dvRegistro
      End If
    End With
  
    '[REVISAR.
    If Len(txtLlave(0).Text) = 1 Then
      MsgBox Choose(gsIdioma, "El Formato debe ser de 2 caracteres.", "The Format must be  2 characters."), vbExclamation
      Cancel = True
      Exit Sub
    End If
    ']

    cmdGrabar.Enabled = True
    upHabilitacion True
    pbValidada = True
   Else
      cmdGrabar.Enabled = False
      upHabilitacion False
      pbValidada = False
   End If
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    lstCOCCoCfg.ListIndex = fila
    lstCOCCoCfg.RemoveItem fila
    lstCOCCoCfg.AddItem codigo & " - " & Trim(txtNombre.Text), fila
  End If
    
End Sub
