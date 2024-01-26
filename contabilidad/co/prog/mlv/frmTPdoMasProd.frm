VERSION 5.00
Begin VB.Form frmTPdoMasProd 
   Caption         =   "[Entidad]"
   ClientHeight    =   3375
   ClientLeft      =   5160
   ClientTop       =   5355
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   7320
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   7
      Left            =   3360
      TabIndex        =   17
      Top             =   2175
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   6
      Left            =   990
      TabIndex        =   16
      Top             =   2175
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   2
      Left            =   990
      TabIndex        =   6
      Top             =   825
      Width           =   6210
   End
   Begin VB.CheckBox chkCalcularIGV 
      Caption         =   "Calcular I.G.&V."
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   5820
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1365
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   4
      Left            =   990
      TabIndex        =   13
      Top             =   1815
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   5
      Left            =   3360
      TabIndex        =   14
      Top             =   1815
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   3
      Left            =   990
      TabIndex        =   8
      Top             =   1140
      Width           =   1000
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   990
      TabIndex        =   4
      Top             =   510
      Width           =   6210
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   135
      Width           =   1440
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   6900
      Picture         =   "frmTPdoMasProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   135
      Width           =   280
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1980
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2670
      Width           =   3480
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
         Picture         =   "frmTPdoMasProd.frx":01AA
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
         Picture         =   "frmTPdoMasProd.frx":02F4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   1220
         Picture         =   "frmTPdoMasProd.frx":03F6
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
         Picture         =   "frmTPdoMasProd.frx":04F8
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   720
      End
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
         Left            =   60
         Picture         =   "frmTPdoMasProd.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   345
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
         Left            =   60
         Picture         =   "frmTPdoMasProd.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Shape shpCuadro 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   15
      Top             =   1485
      Width           =   5235
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   7
      Left            =   60
      TabIndex        =   15
      Top             =   2205
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducción :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   1185
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   525
      Width           =   510
   End
   Begin VB.Label lblDatoDeta 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2430
      TabIndex        =   2
      Top             =   135
      Width           =   4500
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Producto :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Precio :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   6
      Left            =   60
      TabIndex        =   12
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "M.N."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   4
      Left            =   1095
      TabIndex        =   10
      Top             =   1515
      Width           =   300
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "M.E."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   5
      Left            =   3450
      TabIndex        =   11
      Top             =   1515
      Width           =   300
   End
End
Attribute VB_Name = "frmTPdoMasProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
Private pnCta_IndCCo As Integer
Private pcCodCta As String
Private pcCodCCo As String
Private pcTpoMon As String

Private Sub Form_Load()
  pbValidada = False
  Me.KeyPreview = True
  
  With frmTPdoGrd                     'Cambiar Formulario de Grid.
    '[Datos
    txtDato(0).MaxLength = .uorstCoPdoCprProd!codprod.DefinedSize
    txtDato(1).MaxLength = .uorstCoPdoCprProd!gloprod.DefinedSize
    txtDato(2).MaxLength = .uorstCoPdoCprProd!gloprodx.DefinedSize
    txtDato(3).MaxLength = 12
    txtDato(4).MaxLength = 16
    txtDato(5).MaxLength = 16
    txtDato(6).MaxLength = 16
    txtDato(7).MaxLength = 16
    
    txtDato(4).TabIndex = Choose(frmTPdo.cboTpoMon.ListIndex + 1, 13, 14)
    txtDato(5).TabIndex = Choose(frmTPdo.cboTpoMon.ListIndex + 1, 14, 13)
    txtDato(6).TabIndex = Choose(frmTPdo.cboTpoMon.ListIndex + 1, 16, 17)
    txtDato(7).TabIndex = Choose(frmTPdo.cboTpoMon.ListIndex + 1, 17, 16)
    ']
  End With
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  cmdAvanzar.Enabled = (Not pbNuevo)
  cmdRetroceder.Enabled = (Not pbNuevo)
  cmdCorregir.Enabled = (Not pbNuevo)
  upHabilitacion pbNuevo
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(8, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Producto :", "Glosa :", "Traducción :", "Cantidad :", "M.N.", "M.E.", "Precio :", "Importe :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Product :", "Gloss :", "Translation :", "Quantity :", "N.C.", "F.C.", "Price :", "Amount :")
  Next nElemento
  cmdGrabar.Caption = Choose(gsIdioma, "&Aceptar", "&Accept")
  chkCalcularIGV.Caption = Choose(gsIdioma, "Calcular I.G.&V.", "Calculate G.&S.T.")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, False, True, True, aLabel
  ']
End Sub

Private Sub Form_Activate()
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub cmdCorregir_Click()
  cmdRetroceder.Enabled = False
  cmdAvanzar.Enabled = False
  cmdCorregir.Enabled = False
  cmdGrabar.Enabled = True
  cmdDeshacer.Enabled = True
  cmdSalir.Enabled = True
  upHabilitacion True
  
  '[Dato con el foco al corregir.       'Cambiar.
  txtDato((frmTPdo.cboTpoMon.ListIndex + 2)).SetFocus
  ']
End Sub

Private Sub cmdGrabar_Click()
  On Error GoTo Err
  
  If Len(Trim(txtDato(0).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If Len(Trim(txtDato(1).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(1).SetFocus: Exit Sub
  If CDec(txtDato(3).Text) <= 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(3).SetFocus: Exit Sub
  If CDec(txtDato(4).Text) <= 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(4).SetFocus: Exit Sub
  If CDec(txtDato(5).Text) <= 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(5).SetFocus: Exit Sub
  With frmTPdoGrd                     'Cambiar Formulario de Grid.
    frmTPdoGrd.uocnnMain.BeginTrans            'INICIA TRANSACCION.
    If pbNuevo Then
      .uorstCoPdoCprProd.AddNew
    End If
    upDatosDesconectados 0
    With .uorstCoPdoCprProd
      If pbNuevo Then
        !UsrCre = gsAbvUsr
        !FyHCre = Now
      Else
        !UsrMdf = gsAbvUsr
        !FyHMdf = Now
      End If
      .Update
    End With
    frmTPdoGrd.uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    If pbNuevo Then
      .uorstCoPdoCprProd.Requery
      frmTPdoMasGrd.ppDatosGrid
      '[Búsqueda de llave actual.     'Cambiar.
      .uorstCoPdoCprProd.Find "codprod='" & txtDato(0).Text & "'"
      ']
      cmdGrabar.Enabled = False
      upHabilitacion True
      
      upDatosPredeterminados
      '[Dato con el foco al añadir.   'Cambiar.
      txtDato(0).SetFocus
      ']
    Else
      cmdRetroceder.Enabled = True
      cmdAvanzar.Enabled = True
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      upHabilitacion False
    End If
  End With
  Exit Sub
Err:
  gpErrores
  frmTPdoGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Private Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub
Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  '[ARREGLAR: Retrocede si Shift está presionado.
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
  ']ARREGLAR.
End Sub
Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub
Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
  Select Case Index
   Case 3
    ' Verifico cantidad
    If CDec(txtDato(3).Text) <= 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      MsgBox Choose(gsIdioma, "No se ha ingresado cantidad del producto", "Not logged quantity product"), vbCritical
      txtDato(Index).SetFocus
      Exit Sub
    End If
    ' Importe total
    txtDato(6).Text = Format(Round(CDec(txtDato(3).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
    txtDato(7).Text = Format(Round(CDec(txtDato(3).Text) * CDec(txtDato(5).Text), 2), FORMATO_NUM_1)
   Case 4, 5
    ' Convierto importe en cero
    If CDec(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      If Index = 4 And frmTPdo.cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index + 1).Text) * CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 5 And frmTPdo.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index - 1).Text) / CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    ElseIf CDec(txtDato(Index).Text) <> 0 Then
      If Index = 4 And frmTPdo.cboTpoMon.ListIndex = TPOMON_NAC_IND And (frmTPdo.txtDato(5).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index + 1).Text = Format(Round(CDec(txtDato(Index).Text) / CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 5 And frmTPdo.cboTpoMon.ListIndex = TPOMON_EXT_IND And (frmTPdo.txtDato(4).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index - 1).Text = Format(Round(CDec(txtDato(Index).Text) * CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    End If
    ' Importe total
    txtDato(6).Text = Format(Round(CDec(txtDato(3).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
    txtDato(7).Text = Format(Round(CDec(txtDato(3).Text) * CDec(txtDato(5).Text), 2), FORMATO_NUM_1)
  End Select
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  Dim dvRegistroActual As Variant

  'Completa con ceros a la izquierda.
  Select Case Index
   Case 0
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    If lblDatoDeta(Index).Caption <> "" Then
      ' Cuenta repetida
      With frmTPdoGrd.uorstCoPdoCprProd
        If Not (.BOF Or .EOF) And .RecordCount > 0 Then
          dvRegistroActual = .Bookmark
          .MoveFirst
          .Find "codprod='" & txtDato(Index).Text & "'"
          If Not .EOF Then
            MsgBox TEXT_8007, vbExclamation
            If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
            Cancel = True
            Exit Sub
          End If
          .Bookmark = dvRegistroActual
        End If
      End With
      pcCodCta = IIf(IsNull(frmTPdoGrd.uorstCoCprProd!codcta), "", frmTPdoGrd.uorstCoCprProd!codcta)
      pcCodCCo = IIf(pcCodCCo = "", IIf(IsNull(frmTPdoGrd.uorstCoCprProd!codcco_def), "", frmTPdoGrd.uorstCoCprProd!codcco_def), pcCodCCo)
      pcTpoMon = Trim(frmTPdoGrd.uorstCoCprProd!tpomon)
      If pcTpoMon = TPOMON_NAC Then
        txtDato(4).Text = Format(CDec(IIf(CDec(txtDato(4).Text) = 0, frmTPdoGrd.uorstCoCprProd!impcpr, txtDato(4).Text)), FORMATO_NUM_1)
        txtDato(5).Text = Format(CDec(IIf(CDec(txtDato(5).Text) = 0, Round(txtDato(4).Text / CDec(frmTPdo.txtDato(3).Text), 2), txtDato(5).Text)), FORMATO_NUM_1)
      Else
        txtDato(5).Text = Format(CDec(IIf(CDec(txtDato(5).Text) = 0, frmTPdoGrd.uorstCoCprProd!impcpr, txtDato(5).Text)), FORMATO_NUM_1)
        txtDato(4).Text = Format(CDec(IIf(CDec(txtDato(4).Text) = 0, Round(txtDato(5).Text * CDec(frmTPdo.txtDato(3).Text), 2), txtDato(4).Text)), FORMATO_NUM_1)
      End If
      ' Actualizo los datos adicionales
      txtDato(4).Tag = Format(txtDato(4).Text, FORMATO_NUM_1)
      txtDato(5).Tag = Format(txtDato(5).Text, FORMATO_NUM_1)

      ' Habilito controles
      cmdGrabar.Enabled = True
        upHabilitacion True
    Else
      cmdGrabar.Enabled = False
      upHabilitacion False
    End If
   Case 3, 4, 5, 6, 7
    txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_1)
  End Select
  Exit Sub

Err:
  gpErrores
  
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0
    modAyuBus.Prod_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub
Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 0
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmTPdoGrd.uorstCoCprProd
      If .RecordCount > 0 Then .MoveFirst
      .Find "codprod='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detprod), "", !detprod)
      End If
    End With
  End Select
End Function
Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  On Error GoTo Err

  With frmTPdoGrd.uorstCoPdoCprProd    'Cambiar RecordSet.
    If tnFase = 0 Then
      ' Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !mespvs = gsMesAct
        !coddpe = frmTPdo.txtLlave(0).Text
        !pdocpr = frmTPdo.txtLlave(1).Text
        !codprod = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      End If
      ' Datos.
      !codcta = pcCodCta
      !codcco = pcCodCCo
      !gloprod = IIf(txtDato(gsIdioma).Text = "", Null, txtDato(gsIdioma).Text)
      !gloprodx = IIf(txtDato(3 - gsIdioma).Text = "", Null, txtDato(3 - gsIdioma).Text)
      !cantiprod = CDec(txtDato(3).Text)
      !impouni_mn = CDec(txtDato(4).Text)
      !impouni_me = CDec(txtDato(5).Text)
      !impprod_mn = CDec(txtDato(6).Text)
      !impprod_me = CDec(txtDato(7).Text)
      !indigv = IIf(chkCalcularIGV.Value = vbChecked, INDCCO_ACT, INDCCO_INA)
    Else
      ' Llaves.
      txtDato(0).Text = IIf(IsNull(!codprod), "", !codprod)
      ' Datos.
      pcCodCta = !codcta
      pcCodCCo = !codcco
      txtDato(gsIdioma).Text = IIf(IsNull(!gloprod), "", !gloprod)
      txtDato(3 - gsIdioma).Text = IIf(IsNull(!gloprodx), "", !gloprodx)
      txtDato(3).Text = Format(!cantiprod, FORMATO_NUM_1)
      txtDato(4).Text = Format(!impouni_mn, FORMATO_NUM_1)
      txtDato(5).Text = Format(!impouni_me, FORMATO_NUM_1)
      txtDato(6).Text = Format(!impprod_mn, FORMATO_NUM_1)
      txtDato(7).Text = Format(!impprod_me, FORMATO_NUM_1)
      chkCalcularIGV.Value = IIf(!indigv = INDCCO_ACT, vbChecked, vbUnchecked)
      
      txtDato(4).Tag = Format(txtDato(4).Text, FORMATO_NUM_1)
      txtDato(5).Tag = Format(txtDato(5).Text, FORMATO_NUM_1)
      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      ppAyuDet 0
    End If
  End With
      
  Exit Sub
Err:
   gpErrores
   
   Resume
      
End Sub
Public Sub upDatosPredeterminados()    'Cambiar.
  Dim sSentencia As String
  
  txtDato(0).Text = ""
  txtDato(1).Text = frmTPdo.txtDato(1).Text
  txtDato(2).Text = frmTPdo.txtDato(2).Text
  txtDato(3).Text = Format(0, FORMATO_NUM_1)
  txtDato(4).Text = Format(0, FORMATO_NUM_1)
  txtDato(5).Text = Format(0, FORMATO_NUM_1)
  txtDato(6).Text = Format(0, FORMATO_NUM_1)
  txtDato(7).Text = Format(0, FORMATO_NUM_1)
  chkCalcularIGV.Value = vbChecked
  pcCodCCo = frmTPdo.txtDato(9).Text
  
  '[ Obtengo los importes restantes
  If pbNuevo Then
    txtDato(6).Text = Format(CDec(frmTPdo.txtDato(4).Text), FORMATO_NUM_1)
    txtDato(7).Text = Format(CDec(frmTPdo.txtDato(5).Text), FORMATO_NUM_1)
    With frmTPdoGrd
      sSentencia = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impprod_mn), 0), 2) AS ImporteMN, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impprod_me), 0), 2) AS ImporteME "
      sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcopdocprprod "
      Set .porstCancel = .uocnnMain.Execute(sSentencia)
      txtDato(6).Text = Format(CDec(txtDato(6).Text) - .porstCancel!ImporteMN, FORMATO_NUM_1)
      txtDato(7).Text = Format(CDec(txtDato(7).Text) - .porstCancel!ImporteME, FORMATO_NUM_1)
      .porstCancel.Close
    End With
  End If
  txtDato(4).Tag = Format(txtDato(4).Text, FORMATO_NUM_1)
  txtDato(5).Tag = Format(txtDato(5).Text, FORMATO_NUM_1)
  ']
  ' Ayudas.
  lblDatoDeta(0).Caption = ""
End Sub
Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Integer
  
  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      If dnContador = 0 Then
        .Item(dnContador).Enabled = pbNuevo
      Else
        .Item(dnContador).Enabled = tbHabilitar
      End If
    Next dnContador
  End With
  txtDato(6).Enabled = False
  txtDato(7).Enabled = False
  chkCalcularIGV.Enabled = tbHabilitar
  'Ayudas.
  cmdDatoAyud(0).Enabled = pbNuevo
  lblDatoDeta(0).Enabled = pbNuevo
End Sub
'[Propio del formulario.
']
Public Property Get zbNuevo() As Boolean
  zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
  pbNuevo = tbNuevo
End Property
