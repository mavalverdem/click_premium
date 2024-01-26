VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fAbcFormulas 
   Caption         =   "Fórmulas"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   Icon            =   "abcformula.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   10005
   Begin VB.TextBox txtSecuencia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   4170
      TabIndex        =   20
      Text            =   "99"
      Top             =   2085
      Width           =   435
   End
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   8745
      Picture         =   "abcformula.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Deshacer último cambio"
      Top             =   1950
      Width           =   390
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   375
      Left            =   8325
      Picture         =   "abcformula.frx":067E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Deshacer último cambio"
      Top             =   1950
      Width           =   390
   End
   Begin VB.ComboBox cmbProcesos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   165
      Width           =   5055
   End
   Begin VB.CommandButton cmdUndo 
      Height          =   375
      Left            =   9495
      Picture         =   "abcformula.frx":0CF0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Deshacer último cambio"
      Top             =   1950
      Width           =   390
   End
   Begin VB.TextBox txtFormulaPrevious 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4980
      Left            =   9795
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   7395
      Visible         =   0   'False
      Width           =   7050
   End
   Begin VB.TextBox txtConcepto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Nombre de Concepto"
      Top             =   1020
      Width           =   7725
   End
   Begin VB.TextBox txtClase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Clase de Planilla"
      Top             =   705
      Width           =   7725
   End
   Begin VB.ListBox lstValues 
      ForeColor       =   &H00FF0000&
      Height          =   2790
      Left            =   90
      TabIndex        =   9
      Top             =   2415
      Width           =   2655
   End
   Begin VB.TextBox txtFormula 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3945
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   2415
      Width           =   7050
   End
   Begin VB.ListBox lstDesValues 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2940
      Left            =   105
      TabIndex        =   14
      Top             =   2370
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Threed.SSCommand cmdControl 
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1530
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Operador Matem"
   End
   Begin Threed.SSCommand cmdControl 
      Height          =   360
      Index           =   1
      Left            =   1620
      TabIndex        =   4
      Top             =   1530
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Operador Lógico"
   End
   Begin Threed.SSCommand cmdControl 
      Height          =   360
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   1530
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Funciones"
   End
   Begin Threed.SSCommand cmdControl 
      Height          =   360
      Index           =   3
      Left            =   4620
      TabIndex        =   6
      Top             =   1530
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Variables"
   End
   Begin Threed.SSCommand cmdControl 
      Height          =   360
      Index           =   4
      Left            =   6120
      TabIndex        =   7
      Top             =   1530
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Conceptos"
   End
   Begin Threed.SSCommand cmdControl 
      Height          =   360
      Index           =   5
      Left            =   7605
      TabIndex        =   22
      Top             =   1530
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "Valores"
   End
   Begin VB.Label Label1 
      Caption         =   "Secuencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   225
      Left            =   3000
      TabIndex        =   21
      Top             =   2130
      Width           =   1230
   End
   Begin VB.Label lblDescriptionValue 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción de Valor"
      ForeColor       =   &H00800000&
      Height          =   1065
      Left            =   90
      TabIndex        =   13
      Top             =   5280
      Width           =   2640
   End
   Begin VB.Label lblTitle 
      Caption         =   "Operadores Matemáticos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   135
      TabIndex        =   10
      Top             =   2145
      Width           =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   90
      X2              =   9855
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   975
      Width           =   1695
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Clase de Planilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   735
      Width           =   1695
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Definición de Fórmulas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   4125
   End
End
Attribute VB_Name = "fAbcFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gConcepto As String
Public gTipoConcepto As String
Dim lCargaConceptos As Boolean
Dim sProceso As String

Private Sub cmbProcesos_Click()

  txtFormula.Text = ""
  txtSecuencia.Text = ""
  sProceso = Trim(Right(cmbProcesos, 10))
  s_Sql = "SELECT secuencia, formulafun FROM plconceproceso WHERE codcls = '" & ps_ClsPlanilla & "' AND codcpc = '" & gConcepto & "' AND codproce = '" & sProceso & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    txtFormula.Text = IIf(IsNull(porstRecordset!formulafun), "", porstRecordset!formulafun)
    txtSecuencia.Text = porstRecordset!secuencia
    porstRecordset.Close
  End If
  Set porstRecordset = Nothing

End Sub

Private Sub RefreshOperadooresMatematicos()

lblTitle.Caption = "Operadores Matemáticos"
lblDescriptionValue.Caption = ""
With lstValues
    .Clear
    .AddItem "+"
    .AddItem "-"
    .AddItem "*"
    .AddItem "/"
    .AddItem "="
    .AddItem "<"
    .AddItem ">"
    .AddItem ">="
    .AddItem "<="
    .AddItem "<>"
End With

With lstDesValues
    .Clear
    .AddItem "Operador Suma"
    .AddItem "Operador Resta"
    .AddItem "Operador Multiplicación"
    .AddItem "Operador División"
    .AddItem "Operador Igual"
    .AddItem "Operador Mayor que"
    .AddItem "Operador Menor que"
    .AddItem "Operador Mayor o Igual que"
    .AddItem "Operador Menor o Igual que"
    .AddItem "Operador Diferente"
End With

End Sub

Private Sub RefreshOperadooresLogicos()

lblTitle.Caption = "Operadores Lógicos"
lblDescriptionValue.Caption = ""
With lstValues
    .Clear
    .AddItem "AND"
    .AddItem "OR"
    '.AddItem "SI" & Space(200) & "SI (Condición) ENTONCES" & Chr$(13) & Chr$(10) & Space(3) & "<Valor Verdadero>" & Chr$(13) & Chr$(10) & "SINO" & Chr$(13) & Chr$(10) & Space(3) & "<Valor Falso>" & Chr$(13) & Chr$(10) & "FIN SI" & Chr$(13) & Chr$(10)
    '.AddItem "SI" & Space(200) & "SI(<Condición1>,<Operador>,<Condición2>,<Valor Verdadero>,<Valor Falso>)"
    .AddItem "TRUE"
    .AddItem "FALSE"
End With

With lstDesValues
    .Clear
    .AddItem "Operador Y"
    .AddItem "Operador O"
    '.AddItem "Operador Condicional SI(Condición, Verdadero, Falso)"
    .AddItem "Valor Logico para evaluar condiciones que cuyo resultado es Verdadero"
    .AddItem "Valor Logico para evaluar condiciones que cuyo resultado es Falso"
End With

End Sub

Private Sub RefreshFunciones()
Dim rs As New ADODB.Recordset
Dim sSQL As String

lblTitle.Caption = "Funciones"
lblDescriptionValue.Caption = ""

lstValues.Clear
lstDesValues.Clear

sSQL = "SELECT * FROM plvarfunc"
sSQL = sSQL & " WHERE tipo = 'F'"
sSQL = sSQL & " ORDER BY orden"

Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveLast
    rs.MoveFirst
    Do While Not rs.EOF
        lstValues.AddItem rs("nombre") & Space(200) & rs("valor")
        lstDesValues.AddItem rs("descripcion")
        rs.MoveNext
    Loop
    rs.Close
End If

End Sub

Private Sub RefreshVariables()
Dim rs As New ADODB.Recordset
Dim sSQL As String

lblTitle.Caption = "Variables"
lblDescriptionValue.Caption = ""

lstValues.Clear
lstDesValues.Clear

sSQL = "SELECT * FROM plvarfunc"
sSQL = sSQL & " WHERE tipo = 'V'"
sSQL = sSQL & " ORDER BY orden"

Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveLast
    rs.MoveFirst
    Do While Not rs.EOF
        lstValues.AddItem rs("nombre")
        lstDesValues.AddItem rs("descripcion")
        rs.MoveNext
    Loop
    rs.Close
End If

End Sub

Private Sub RefreshValores()
Dim rs As New ADODB.Recordset
Dim sSQL As String

lblTitle.Caption = "Valores"
lblDescriptionValue.Caption = ""

lstValues.Clear
lstDesValues.Clear

sSQL = "SELECT * FROM pltablabase"
sSQL = sSQL & " WHERE codcls = '" & ps_ClsPlanilla & "'"
sSQL = sSQL & " AND pdoano = '" & ps_Anyo & "'"
sSQL = sSQL & " ORDER BY codtbl"

Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveLast
    rs.MoveFirst
    Do While Not rs.EOF
        lstValues.AddItem rs("destbl") & Space(200) & "K_" & rs("codtbl")
        lstDesValues.AddItem "Código " & rs("codtbl")
        rs.MoveNext
    Loop
    rs.Close
End If

End Sub

Private Sub RefreshConceptos()

Dim rs As New ADODB.Recordset
Dim sSQL As String

lblTitle.Caption = "Conceptos"
lblDescriptionValue.Caption = ""
lstValues.Clear
lstDesValues.Clear

sSQL = "SELECT cxc.codcpc, cpc.descpc "
sSQL = sSQL & " FROM plconceproceso cxc, plconcepto cpc "
sSQL = sSQL & " WHERE cxc.codcls='" & ps_ClsPlanilla & "'"
sSQL = sSQL & " AND cxc.codproce='" & Trim(Right(cmbProcesos.Text, 2)) & "'"
sSQL = sSQL & " AND cpc.codcpc=cxc.codcpc"
sSQL = sSQL & " AND cpc.codcpc<>'" & gConcepto & "'"
sSQL = sSQL & " ORDER BY cxc.secuencia"

Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveLast
    rs.MoveFirst
    Do While Not rs.EOF
        lstValues.AddItem rs("descpc") & Space(200) & "C" & rs("codcpc")
        lstDesValues.AddItem "Concepto " & rs("codcpc")
        rs.MoveNext
    Loop
    rs.Close
End If

End Sub

Private Sub cmdControl_Click(Index As Integer)
  
  Select Case Index
   Case 0: Call RefreshOperadooresMatematicos
   Case 1: Call RefreshOperadooresLogicos
   Case 2: Call RefreshFunciones
   Case 3: Call RefreshVariables
   Case 4: Call RefreshConceptos
   Case 5: Call RefreshValores
  End Select

End Sub

Private Sub cmdExit_Click()

If MsgBox("Seguro de salir sin grabar Formula?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Planillas") <> vbYes Then
    txtFormula.SetFocus
    Exit Sub
End If

Unload Me

End Sub

Private Sub cmdUndo_Click()

cmdUndo.Enabled = False
txtFormula.Text = txtFormulaPrevious.Text
txtFormula.SetFocus

End Sub

Private Sub cmdUpdate_Click()
  Dim nOcurrencias As Integer
  Dim nSecuencia As Long
  Dim dFecha As String
  
  If MsgBox("Seguro de grabar Formula?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sistema de Planillas") <> vbYes Then
    txtFormula.SetFocus
    Exit Sub
  End If

  s_Sql = "SELECT COUNT(*) As nOcurrencias"
  s_Sql = s_Sql & " FROM plconceproceso cxp"
  s_Sql = s_Sql & " WHERE cxp.codcls='" & ps_ClsPlanilla & "'"
  s_Sql = s_Sql & " AND cxp.codproce='" & sProceso & "'"
  s_Sql = s_Sql & " AND cxp.codcpc='" & gConcepto & "'"
  nOcurrencias = 0
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    nOcurrencias = porstRecordset!nOcurrencias
  End If
  porstRecordset.Close

  If nOcurrencias = 0 Then
    s_Sql = "SELECT MAX(secuencia) As nSecuencia"
    s_Sql = s_Sql & " FROM plconceproceso cxp"
    s_Sql = s_Sql & " WHERE cxp.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & " AND cxp.codproce='" & sProceso & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      nSecuencia = IIf(IsNull(porstRecordset!nSecuencia), 0, porstRecordset!nSecuencia)
      nSecuencia = nSecuencia + 1
      porstRecordset.Close
    End If
    dFecha = Format(Now, "yyyy-mm-dd")
    s_Sql = "INSERT INTO plconceproceso (codcls,codproce,codcpc,secuencia,usrcre,fyhcre,usrmdf,fyhmdf,formulafun) values('" & ps_ClsPlanilla & "','" & sProceso & "','" & gConcepto & "'," & nSecuencia & ",'" & ps_UserId & "','" & dFecha & "',NULL,NULL,'" & txtFormula.Text & "')"
  Else
    s_Sql = "UPDATE plconceproceso cxp"
    s_Sql = s_Sql & " SET cxp.formulafun = '" & txtFormula.Text & "', "
    s_Sql = s_Sql & " cxp.secuencia = '" & txtSecuencia.Text & "'"
    s_Sql = s_Sql & " WHERE cxp.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & " AND cxp.codproce='" & sProceso & "'"
    s_Sql = s_Sql & " AND cxp.codcpc='" & gConcepto & "'"
  End If
  'Ejecuta Instrucción
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  Set porstRecordset = Nothing

End Sub

Private Sub Form_Load()

Dim nTop As Integer, nLeft As Integer, nHeight As Integer, nWidth As Integer
Dim rs As New ADODB.Recordset
Dim sSQL As String

nHeight = 6930
nLeft = 1000
nTop = 80
nWidth = 10155

With Me
    .Height = nHeight
    .Left = nLeft
    .Top = nTop
    .Width = nWidth
End With

lblDescriptionValue.Caption = ""
txtFormula.Text = ""
txtFormula.Locked = (gTipoConcepto <> "F")

txtFormulaPrevious.Text = ""
txtFormulaPrevious.Visible = False
cmdUndo.Enabled = False

txtClase.Text = ""
txtConcepto.Text = ""
cmbProcesos.Clear

sSQL = "SELECT codcls, descls "
sSQL = sSQL & "FROM plclasplan "
sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
sSQL = sSQL & "AND estadocls='" & s_Estado_Act & "'"
Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
  txtClase.Text = rs("descls")
  rs.Close
End If

sSQL = "SELECT codcpc, descpc "
sSQL = sSQL & "FROM plconcepto "
sSQL = sSQL & "WHERE codcpc='" & gConcepto & "' "
sSQL = sSQL & "AND estadocpc='" & s_Estado_Act & "'"
Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
  txtConcepto.Text = gConcepto & " - " & rs("descpc")
  rs.Close
End If

sSQL = "SELECT codcls, codproce, desproce "
sSQL = sSQL & "FROM plproceso "
sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
sSQL = sSQL & "AND estadoproce<>'" & s_EstadoRemAper & "'"
Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
  Do While Not rs.EOF
    cmbProcesos.AddItem rs("desproce") & Space(200) & rs("codproce")
    rs.MoveNext
  Loop
  rs.Close
  cmbProcesos.ListIndex = fFormulaConcepto.cmbProcesos.ListIndex
End If

Call RefreshOperadooresMatematicos

End Sub

Private Sub lstValues_Click()

lstDesValues.ListIndex = lstValues.ListIndex
lblDescriptionValue.Caption = lstDesValues.Text

End Sub

Private Sub lstValues_DblClick()

Dim sValor As String

If (gTipoConcepto <> "F") Then
  Exit Sub
End If

sValor = Mid(lstValues, 200)
If sValor = "" Then
    sValor = Mid(lstValues, 1, 200)
End If

txtFormulaPrevious.Text = txtFormula.Text

txtFormula.Text = Mid(txtFormula.Text, 1, txtFormula.SelStart) & Trim(sValor) & Mid(txtFormula.Text, txtFormula.SelStart + 1)
cmdUndo.Enabled = True

End Sub
