VERSION 5.00
Begin VB.Form frmMCCo 
   Appearance      =   0  'Flat
   Caption         =   "[Entidad]"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   6885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIndPdo 
      Caption         =   "Pedido de Compra"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1020
      TabIndex        =   15
      Top             =   1515
      Width           =   2700
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
      Index           =   1
      Left            =   1020
      TabIndex        =   5
      Top             =   1095
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   60
      TabIndex        =   14
      Top             =   1740
      Width           =   975
      Begin VB.CheckBox chkEstCCo 
         Caption         =   "Activo"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   165
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1702
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1815
      Width           =   3480
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
         Picture         =   "frmMCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   360
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
         Picture         =   "frmMCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   338
         Width           =   360
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
         Picture         =   "frmMCCo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmMCCo.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmMCCo.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   2690
         Picture         =   "frmMCCo.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   12
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
      Left            =   1020
      TabIndex        =   3
      Top             =   720
      Width           =   4335
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
      Left            =   1335
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducción:"
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
      Left            =   60
      TabIndex        =   4
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
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
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Centro de Costo:"
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
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   6840
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMCCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmMCCoGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codcco.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
      txtDato(gsIdioma - 1).MaxLength = .uorstMain!detcco.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain!DetCCox.DefinedSize
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Centro de Costo:", "Descripción:", "Traducción:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Cost Center:", "Description:", "Translation:")
  Next nElemento
  chkIndPdo.Caption = Choose(gsIdioma, "Pedido de Compra", "Order of Purchase")
  chkEstCCo.Caption = Choose(gsIdioma, "&Activo", "&Active")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']
   
End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
'   If txtDato(0).Text <> "" Then ppAyuDet 0
 ']

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMCCoGrd.uorstMain.BOF And frmMCCoGrd.uorstMain.EOF) Then
   frmMCCoGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMCCoGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMCCoGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmMCCoGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
'            !FyHMdf = Now
         End If
         .Update
      End With
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain.Find "CodCCo='" & txtLlave(0).Text & "'"
       ']
         cmdGrabar.Enabled = False
         upHabilitacion False
   
         upDatosPredeterminados
       '[Llave con el foco al añadir.  'Cambiar.
         txtLlave(0).SetFocus
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
  
   frmMCCoGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
'   Select Case Index                   'Cambiar. Añadir índices.
'   Case 0, 1
'      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
'   End Select
'   ppAyuBus Index
End Sub

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

Private Sub txtllave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
'         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
'      End If
'   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMCCoGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodCCo='" & txtLlave(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistro <> -1 Then .Bookmark = dvRegistro
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistro
         End If
      End With
      
'[REVISAR Miguel Angel
  Dim n_Index, n_Nivel As Integer
  Dim a_Nivel(4) As Integer
  
  cmdGrabar.Enabled = True
  upHabilitacion True
  pbValidada = True
  
  frmMCCoGrd.uorstCodCCo.Requery '2016-01-31 refresca recorset
  
   For n_Nivel = 1 To Len(gsNivCCo)
      a_Nivel(n_Nivel) = CInt(Mid(gsNivCCo, n_Nivel, 1))
      If Len(txtLlave(Index).Text) = a_Nivel(n_Nivel) Then
        With frmMCCoGrd.uorstCodCCo
          For n_Index = n_Nivel To 1 Step -1
             If n_Index = 1 Then Exit Sub
             .MoveFirst
             .Find "CodCCo = '" & Left(txtLlave(Index).Text, a_Nivel(n_Index - 1)) & "'"
             If .EOF Then
                MsgBox Choose(gsIdioma, "El Centro de costo ", "The cost center ") & Left(txtLlave(Index).Text, a_Nivel(n_Index - 1)) & Choose(gsIdioma, " (nivel anterior) NO EXISTE.", " (previous level) NOT EXIST."), vbInformation + vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
                Cancel = True
                Exit Sub
             End If
          Next n_Index
        End With
      End If
      If n_Nivel = Len(gsNivCCo) Then
        MsgBox Choose(gsIdioma, "El Nivel ", "The Level ") & Str(Len(txtLlave(Index).Text)) & Choose(gsIdioma, " no está configurado.", "is not configurated"), vbInformation + vbOKOnly, Choose(gsIdioma, "Advertencia", "Warning")
        Cancel = True
        Exit Sub
      End If
   Next n_Nivel
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

 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
'   On Error GoTo Err

  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 12                             'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

  'Asigna 0 a campos numéricos si están vacíos.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select

  'Busca el dato en su tabla principal.
'   Select Case Index
'   Case 12                             'Cambiar (añadir índices).
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
      
'   Exit Sub
'Err:
'   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
'   Select Case tnIndex
'   Case 2                              'Cambiar (añadir índices).
'      modAyuBus.Dtt_Cod txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
'      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
'   Select Case tnIndex                 'Cambiar.
'   Case 2
'      If txtDato(tnIndex).Text = "" Then
'         lblDatoDeta(tnIndex).Caption = ""
'         Exit Function
'      End If
'      With porstTGDtt
'         .MoveFirst
'         .Find "CodDtt='" & txtDato(tnIndex).Text & "'"
'         If .EOF Then
'            MsgBox TEXT_8006, vbExclamation
'            ppAyuDet = True
'         Else
'            lblDatoDeta(tnIndex).Caption = " " & !DetDtt
'         End If
'      End With
'   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
  On Error GoTo Err
  
  With frmMCCoGrd.uorstMain           'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !codcco = txtLlave(0).Text
      End If
      'Datos.
      !detcco = txtDato(gsIdioma - 1).Text
      !DetCCox = txtDato(2 - gsIdioma).Text
      !indpdocpr = IIf(chkIndPdo.Value = vbChecked, INDCCO_ACT, INDCCO_INA)
      !EstCCo = IIf(chkEstCCo.Value = vbChecked, ESTCCO_ACT, ESTCCO_INA)
    Else
      'Llaves.
      txtLlave(0).Text = !codcco
      'Datos.
      txtDato(gsIdioma - 1).Text = IIf(IsNull(!detcco), "", !detcco)
      txtDato(2 - gsIdioma).Text = IIf(IsNull(!DetCCox), "", !DetCCox)
      chkIndPdo.Value = IIf(!indpdocpr = INDCCO_ACT, vbChecked, vbUnchecked)
      chkEstCCo.Value = IIf(!EstCCo = ESTCCO_ACT, vbChecked, vbUnchecked)
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
   chkEstCCo.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With

  'Ayudas.
'   lblDatoDeta(2).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   chkEstCCo.Enabled = tbHabilitar
   chkIndPdo.Enabled = tbHabilitar
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
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

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdCorregir.Enabled = IIf(pbNuevo, False, taOpciones(0))
End Property


