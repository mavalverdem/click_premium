VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOSelPen 
   Caption         =   "[Entidad]"
   ClientHeight    =   4995
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   6945
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   6945
   Visible         =   0   'False
   Begin VB.TextBox txtCelda 
      Enabled         =   0   'False
      Height          =   280
      Left            =   -15
      TabIndex        =   5
      Top             =   3375
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6945
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6945
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
         Left            =   5430
         Picture         =   "frmOSelPen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
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
         Left            =   6180
         Picture         =   "frmOSelPen.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
      Begin VB.Frame fraBuscar 
         Caption         =   "&Buscar por [Columna]"
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
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   200
            Width           =   2415
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgPendiente 
      Bindings        =   "frmOSelPen.frx":024C
      Height          =   4365
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   7699
      _Version        =   393216
      BackColorFixed  =   16777152
      ForeColorFixed  =   16711680
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      Appearance      =   0
   End
End
Attribute VB_Name = "frmOSelPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public usConnStrgSele As String, usConnStrgOrde As String
Public uorstMain As ADODB.Recordset
Public unArribaFormulario As Integer, _
       unIzquierdaFormulario As Integer, _
       unAltoFormulario As Integer, _
       unAnchoFormulario As Integer
Public unElementos As Integer, uaWhere As Integer
Public uaTitulos As Variant, uaAncho As Variant, _
       uaFormato As Variant, uaAlineamiento As Variant, _
       uaOrden As Variant
Public uvDato1Posicion As Integer, uvDato1Previo As Variant, uvDato1 As Variant
Public uvDato2Posicion As Integer, uvDato2 As Variant
Public usCriterio As String
Private pnColumnaOrd As Integer
Private s_Primary As String
'[
Private Sub ppCeldaGrilla()
  If Not IsNumeric(txtCelda.Text) Then: MsgBox Choose(gsIdioma, "Ingresó Incorrecto, Los Valores deben ser numéricos", "Information isn't valid, The values must be numerics"), vbCritical, "Importe de Celda": Exit Sub
  If mfgPendiente.Col >= 4 And mfgPendiente.Col <= 5 Then
    mfgPendiente.TextMatrix(mfgPendiente.Row, mfgPendiente.Col) = FormatNumber(CDec(txtCelda.Text), 2)
    mfgPendiente = FormatNumber(CDec(txtCelda.Text), 2)
  ElseIf mfgPendiente.Col = 6 Then
    mfgPendiente.TextMatrix(mfgPendiente.Row, mfgPendiente.Col) = Trim(txtCelda.Text)
  End If
  ppCeldaRecordset mfgPendiente.Row, mfgPendiente.Col
End Sub

Private Sub ppCeldaRecordset(n_Fila As Long, n_Columna As Long)
  Dim n_ImporteMN As Double, n_ImporteME As Double
  Dim s_CenCosto As String, s_Seleccion As String
  Dim sSentencia As String
  
  ' Ubico la cuenta a actualizar
  uorstMain.MoveFirst
  uorstMain.Find "cLlave='" & s_Primary & "'"
  If uorstMain.EOF Then
    MsgBox Choose(gsIdioma, "Celda no actualizable", "This Cell can not be up to date"), vbInformation
    Exit Sub
  End If
  n_ImporteMN = CDec(uorstMain!imppmn)
  n_ImporteME = CDec(uorstMain!imppme)
  s_CenCosto = IIf(IsNull(uorstMain!codcco), "", uorstMain!codcco)
  s_Seleccion = uorstMain!indsel
  Select Case n_Columna
   Case 4
    n_ImporteMN = Abs(CDec(mfgPendiente.TextMatrix(n_Fila, n_Columna)))
    n_ImporteMN = n_ImporteMN * IIf(uorstMain!cImpSaldo < 1, -1, 1)
    n_ImporteME = IIf(uorstMain!tpomon = TPOMON_NAC, Round(n_ImporteMN / CDec(frmTBanCab.txtDato(6).Text), 2), n_ImporteME)
   Case 5
    n_ImporteME = Abs(CDec(mfgPendiente.TextMatrix(n_Fila, n_Columna)))
    n_ImporteME = n_ImporteME * IIf(uorstMain!cImpSaldo < 1, -1, 1)
    n_ImporteMN = IIf(uorstMain!tpomon = TPOMON_EXT, Round(n_ImporteME * CDec(frmTBanCab.txtDato(6).Text), 2), n_ImporteMN)
   Case 6
    s_CenCosto = Trim(mfgPendiente.TextMatrix(n_Fila, n_Columna))
    ' Convierto al tipo de cambio
    n_ImporteMN = CDec(mfgPendiente.TextMatrix(n_Fila, 4))
    n_ImporteME = CDec(mfgPendiente.TextMatrix(n_Fila, 5))
   Case 7
    s_Seleccion = IIf(mfgPendiente.TextMatrix(n_Fila, n_Columna) = "Si", INDPREGEN_ACT, INDPREGEN_INA)
    ' Convierto al tipo de cambio
    n_ImporteMN = CDec(mfgPendiente.TextMatrix(n_Fila, 4))
    n_ImporteME = CDec(mfgPendiente.TextMatrix(n_Fila, 5))
  End Select
  
  ' Actualizo el importe de la cuenta
  frmTBanGrd.uocnnMain.BeginTrans            'INICIA TRANSACCION.
  sSentencia = "UPDATE codoctmp1 SET "
  sSentencia = sSentencia & "imppmn=" & n_ImporteMN & ", "
  sSentencia = sSentencia & "imppme=" & n_ImporteME & ", "
  sSentencia = sSentencia & "codcco='" & IIf(s_CenCosto = "", Null, s_CenCosto) & "', "
  sSentencia = sSentencia & "indsel ='" & s_Seleccion & "' "
  sSentencia = sSentencia & "WHERE " & IIf(ps_Plataforma = pSrvMySql, "Concat(codcta,codtdc,serdoc,nrodoc)", "(codcta+codtdc+serdoc+nrodoc)") & "='" & s_Primary & "'"
  frmTBanGrd.uocnnMain.Execute sSentencia
  frmTBanGrd.uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  ' Actualizo la fila de la grilla
  uorstMain.Requery
  ppFilaGrilla n_Fila
            
End Sub

Private Sub ppEditGrilla(o_Grilla As MSFlexGrid, o_Texto As TextBox, n_KeyAscii As Integer)

  ' Utiliza el carácter escrito
  Select Case n_KeyAscii
   Case 0 To vbKeySpace   ' Un espacio significa modificar el texto actual
    o_Texto = o_Grilla
    o_Texto.SelStart = 1000
   Case Else              ' Otro carácter reemplaza el texto actual
    If Not IsNumeric(o_Texto.Text) Then: MsgBox Choose(gsIdioma, "Ingresó Incorrecto, Los Valores deben ser numéricos", "Information isn't valid, The values must be numerics"), vbCritical, "Importe de Celda": Exit Sub
    o_Texto.SelStart = 1
  End Select
  ' Muestra la celda en la posición correcta
  s_Primary = o_Grilla.TextMatrix(o_Grilla.Row, 0) & o_Grilla.TextMatrix(o_Grilla.Row, 1)
  o_Texto.Move o_Grilla.Left + o_Grilla.CellLeft, o_Grilla.Top + o_Grilla.CellTop, o_Grilla.CellWidth, o_Grilla.CellHeight
  o_Texto.Visible = True
  o_Texto.Enabled = True
  o_Texto.SetFocus

End Sub

Private Sub ppEditKeyCode(o_Grilla As MSFlexGrid, o_Texto As TextBox, n_KeyCode As Integer)
    
  Select Case n_KeyCode
   Case vbKeyEscape, vbKeyReturn  ' ESC, ENTER: ocultar, devuelve el enfoque a la grilla
    o_Texto.Visible = False
    o_Grilla.SetFocus
   Case vbKeyUp                   ' Arriba: ocultar, devuelve el enfoque a la grilla
    o_Texto.Visible = False
    If o_Grilla.Row > o_Grilla.FixedRows Then
        o_Grilla.Row = o_Grilla.Row - 1
    End If
    o_Grilla.SetFocus
   Case vbKeyDown                 ' Abajo: ocultar, devuelve el enfoque a la grilla
    o_Texto.Visible = False
    If o_Grilla.Row < (o_Grilla.Rows - 1) Then
        o_Grilla.Row = o_Grilla.Row + 1
    End If
    o_Grilla.SetFocus
  End Select

End Sub

Private Sub ppFilaGrilla(n_Fila As Long)
  Dim n_aImporte(3) As Double
  
   ' Ubico la cuenta a actualizar
  uorstMain.MoveFirst
  uorstMain.Find "cLlave='" & s_Primary & "'"
  If uorstMain.EOF Then
    MsgBox Choose(gsIdioma, "Celda no actualizable", "This Cell can not be up to date"), vbInformation
    Exit Sub
  End If
 ' Obtengo los importes iniciales
  n_aImporte(1) = CDec(uorstMain!cImpSaldo)
  n_aImporte(2) = CDec(uorstMain!imppmn)
  n_aImporte(3) = CDec(uorstMain!imppme)
  ' Fila modificable de la grilla
  With mfgPendiente
    .TextMatrix(n_Fila, 0) = uorstMain!codcta
    .TextMatrix(n_Fila, 1) = IIf(IsNull(uorstMain!cDocume), "", uorstMain!cDocume)
    .TextMatrix(n_Fila, 2) = IIf(IsNull(uorstMain!cTpoMon), "", uorstMain!cTpoMon)
    .TextMatrix(n_Fila, 3) = FormatNumber(n_aImporte(1), 2)
    .TextMatrix(n_Fila, 4) = FormatNumber(n_aImporte(2), 2)
    .TextMatrix(n_Fila, 5) = FormatNumber(n_aImporte(3), 2)
    .TextMatrix(n_Fila, 6) = IIf(IsNull(uorstMain!codcco), "", uorstMain!codcco)
    .TextMatrix(n_Fila, 7) = IIf(uorstMain!indsel = INDPREGEN_ACT, "Si", "No")
  End With

End Sub

Private Sub ppInicializaGrilla()
  Dim n_Index As Integer
    
  With mfgPendiente
    .Cols = 8
    .FixedCols = 4
    .Rows = 2
    .FixedRows = 1
    .GridColor = vbRed
    .GridColorFixed = vbBlue
    .GridLines = flexGridFlat
    .GridLinesFixed = flexGridInset
    .GridLineWidth = 1
    .SelectionMode = flexSelectionFree
    .BackColor = &H80000018
    .BackColorBkg = &H8000000F
    .BackColorFixed = &HFFFFC0
    .BackColorSel = &H8000000D
    .TextStyleFixed = flexTextRaisedLight
    .ForeColor = vbBlack
    .ForeColorFixed = vbBlue
    .FillStyle = flexFillSingle
  End With
    
  For n_Index = 0 To (mfgPendiente.Cols - 1)
    mfgPendiente.Col = n_Index
    mfgPendiente.TextMatrix(0, n_Index) = uaTitulos(n_Index)
    mfgPendiente.ColAlignment(n_Index) = uaAlineamiento(n_Index)
    mfgPendiente.ColWidth(n_Index) = uaAncho(n_Index)
  Next n_Index
  ' Incializo la altura de la fila inicial
  mfgPendiente.RowHeight(1) = 0
  
End Sub

Private Sub ppRegistrosGrilla()
  Dim n_Index As Integer
  Dim n_aImporte(3) As Double
  
  mfgPendiente.Redraw = False
  ' Elimino y configuro la grilla
  mfgPendiente.Clear
  ppInicializaGrilla
  n_Index = 1
  uorstMain.Requery
  If uorstMain.RecordCount > 0 Then
    uorstMain.MoveFirst
    n_Index = 3
    Do While Not uorstMain.EOF
      ' Obtengo los importes iniciales
      n_aImporte(1) = CDec(uorstMain!cImpSaldo)
      n_aImporte(2) = CDec(IIf(uorstMain!tpomon = TPOMON_NAC, uorstMain!cImpSaldo, Round(uorstMain!cImpSaldo * CDec(frmTBanCab.txtDato(6)), 2)))
      n_aImporte(3) = CDec(IIf(uorstMain!tpomon = TPOMON_EXT, uorstMain!cImpSaldo, Round(uorstMain!cImpSaldo / CDec(frmTBanCab.txtDato(6)), 2)))
      With mfgPendiente
        .Rows = n_Index
        .TextMatrix(n_Index - 1, 0) = uorstMain!codcta
        .TextMatrix(n_Index - 1, 1) = IIf(IsNull(uorstMain!cDocume), "", uorstMain!cDocume)
        .TextMatrix(n_Index - 1, 2) = IIf(IsNull(uorstMain!cTpoMon), "", uorstMain!cTpoMon)
        .TextMatrix(n_Index - 1, 3) = FormatNumber(n_aImporte(1), 2)
        .TextMatrix(n_Index - 1, 4) = FormatNumber(n_aImporte(2), 2)
        .TextMatrix(n_Index - 1, 5) = FormatNumber(n_aImporte(3), 2)
        .TextMatrix(n_Index - 1, 6) = IIf(IsNull(uorstMain!codcco), "", uorstMain!codcco)
        .TextMatrix(n_Index - 1, 7) = IIf(uorstMain!indsel = INDPREGEN_ACT, "Si", "No")
      End With
      ' Incremento las filas
      n_Index = n_Index + 1
      uorstMain.MoveNext
    Loop
  End If
  mfgPendiente.Redraw = True
  
End Sub

Private Sub Form_Load()
  Me.Top = unArribaFormulario
  Me.Left = unIzquierdaFormulario
  Me.Height = unAltoFormulario
  Me.Width = unAnchoFormulario
 
 '[Recordsets                          'Cambiar.
  Set uorstMain = New ADODB.Recordset
  With uorstMain
    .ActiveConnection = frmTBanGrd.uocnnMain
    .Source = usConnStrgSele & usConnStrgOrde
'    .CursorLocation = adUseClient 'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "codoctmp1"
    If .RecordCount <> 0 Then
      .Find usCriterio
      If .EOF Then .MoveFirst
    End If
  End With
 ']
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  Me.Caption = Choose(gsIdioma, "Ayuda", "Help")
  CaptionBotones Me, True, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
  ppRegistrosGrilla
  ' Configura el texto de la celda
  txtCelda.Text = 0
  txtCelda.Alignment = vbRightJustify
  txtCelda.MaxLength = 18
  txtCelda.Enabled = False
  txtCelda.Visible = False
  
End Sub

Private Sub Form_Activate()
  fraBuscar.Caption = TEXT_BUSCA & mfgPendiente.TextMatrix(0, 1)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   cmdSalir.Left = Me.Width - 840
   cmdAceptar.Left = cmdSalir.Left - 790
   fraBuscar.Width = cmdAceptar.Left - fraBuscar.Left - 50
   txtBuscar.Width = fraBuscar.Width - 240
   mfgPendiente.Height = Me.Height - 450 - picOpciones.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   uorstMain.Close
   Set uorstMain = Nothing
End Sub

Private Sub cmdAceptar_Click()
  On Error GoTo Err
  
  uaWhere = INDPREGEN_ACT
  uvDato1 = mfgPendiente.TextMatrix(mfgPendiente.Row, 1)
  uvDato2 = mfgPendiente.TextMatrix(mfgPendiente.Row, 3)
  Unload Me
  
  Exit Sub
Err:
  If Err.Number = 13 Then  '13=El tipo no concide. Aparece si uvDato2 no tiene valor.
    Resume Next
  Else
    MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  End If
End Sub

Private Sub cmdSalir_Click()
  uaWhere = INDPREGEN_INA
  uvDato1 = uvDato1Previo
  Unload Me
End Sub

Private Sub mfgPendiente_DblClick()
  ' Verificar registros
  If mfgPendiente.Rows = 2 Or mfgPendiente.Row = 1 Then Exit Sub
  If Not (mfgPendiente.Col >= 4 And mfgPendiente.Col <= 7) Then MsgBox Choose(gsIdioma, "Columna no actualizable", "This Col can not be up to date"), vbInformation: Exit Sub
  If mfgPendiente.Col >= 4 And mfgPendiente.Col <= 6 Then
    ppEditGrilla mfgPendiente, txtCelda, vbKeySpace ' Simula un espacio
  End If
  If mfgPendiente.Col = 7 Then
    mfgPendiente.TextMatrix(mfgPendiente.Row, mfgPendiente.Col) = IIf(mfgPendiente.TextMatrix(mfgPendiente.Row, mfgPendiente.Col) = "Si", "No", "Si")
    s_Primary = mfgPendiente.TextMatrix(mfgPendiente.Row, 0) & mfgPendiente.TextMatrix(mfgPendiente.Row, 1)
    ppCeldaRecordset mfgPendiente.Row, mfgPendiente.Col
  End If
End Sub
Private Sub mfgPendiente_KeyPress(KeyAscii As Integer)
  ' Verificar registros
  If mfgPendiente.Rows = 2 Or mfgPendiente.Row = 1 Then Exit Sub
  If Not (mfgPendiente.Col >= 4 And mfgPendiente.Col <= 6) Then MsgBox Choose(gsIdioma, "Columna no actualizable", "This Col can not be up to date"), vbInformation: Exit Sub
  ppEditGrilla mfgPendiente, txtCelda, KeyAscii
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With uorstMain
    dvRegistroActual = .Bookmark
    dsCriterio = "cDocume LIKE '" & Trim(txtBuscar) & "*'"
    .Find dsCriterio, , , 1
    If .EOF = True Then
      .Bookmark = dvRegistroActual
    End If
  End With
   Exit Sub
Err:
  If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
     uorstMain.Bookmark = dvRegistroActual
  Else
     MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  End If
End Sub

Private Sub txtCelda_GotFocus()
  txtCelda.SelStart = 0
  txtCelda.SelLength = txtCelda.MaxLength
End Sub
Private Sub txtCelda_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub txtCelda_KeyUp(KeyCode As Integer, Shift As Integer)
  ppCeldaGrilla
  ppEditKeyCode mfgPendiente, txtCelda, KeyCode
End Sub
Private Sub txtCelda_LostFocus()
  txtCelda.Text = 0
  txtCelda.Enabled = False
  txtCelda.Visible = False
End Sub
Private Sub txtCelda_Validate(Cancel As Boolean)
  Dim sMoneda As String
  Dim nImporte As Double
  
  txtCelda.Text = IIf(Not IsNumeric(txtCelda.Text), 0, txtCelda.Text)
  If mfgPendiente.Col >= 4 And mfgPendiente.Col <= 5 Then
    sMoneda = mfgPendiente.TextMatrix(mfgPendiente.Row, 2)
    nImporte = CDec(mfgPendiente.TextMatrix(mfgPendiente.Row, 3))
    If (sMoneda = "S/." And mfgPendiente.Col = 4 And Abs(nImporte) < CDec(txtCelda.Text)) Then MsgBox Choose(gsIdioma, "El importe es mayor al saldo", "The amount is more than saldo"), vbCritical, "Importe de Celda": Cancel = True: Exit Sub
    If (sMoneda = "US$" And mfgPendiente.Col = 5 And Abs(nImporte) < Abs(CDec(txtCelda.Text))) Then MsgBox Choose(gsIdioma, "El importe es mayor al saldo", "The amount is more than saldo"), vbCritical, "Importe de Celda": Cancel = True: Exit Sub
    txtCelda.Text = IIf(CDec(txtCelda.Text) < 0, 0, txtCelda.Text)
    txtCelda.Text = FormatNumber(txtCelda.Text, 2)
  End If
End Sub

Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property
