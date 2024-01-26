VERSION 5.00
Begin VB.Form frmMAsiDeta 
   Caption         =   "[Entidad]"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   975
      TabIndex        =   12
      Top             =   1695
      Width           =   980
   End
   Begin VB.TextBox txtLlave 
      Height          =   300
      Index           =   0
      Left            =   975
      TabIndex        =   3
      Top             =   465
      Width           =   980
   End
   Begin VB.TextBox txtLlave 
      Height          =   300
      Index           =   1
      Left            =   975
      TabIndex        =   6
      Top             =   795
      Width           =   980
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   6135
      Picture         =   "frmMAsiDeta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1335
      Width           =   255
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
      Left            =   975
      TabIndex        =   9
      Top             =   1335
      Width           =   980
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   6135
      Picture         =   "frmMAsiDeta.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   795
      Width           =   255
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   6135
      Picture         =   "frmMAsiDeta.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   465
      Width           =   255
   End
   Begin VB.ComboBox cboColumna 
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "frmMAsiDeta.frx":04FE
      Left            =   975
      List            =   "frmMAsiDeta.frx":0500
      TabIndex        =   1
      Top             =   90
      Width           =   2970
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1462
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
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
         Picture         =   "frmMAsiDeta.frx":0502
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmMAsiDeta.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmMAsiDeta.frx":074E
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmMAsiDeta.frx":0850
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmMAsiDeta.frx":099A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   338
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
         Picture         =   "frmMAsiDeta.frx":0B44
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   360
      End
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
      Height          =   300
      Index           =   0
      Left            =   1950
      TabIndex        =   10
      Top             =   1335
      Width           =   4185
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "C.Costo :"
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
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   1380
      Width           =   660
   End
   Begin VB.Label lblLlaveDeta 
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
      Height          =   300
      Index           =   1
      Left            =   1950
      TabIndex        =   7
      Top             =   795
      Width           =   4185
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta ME :"
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
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblLlaveDeta 
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
      Height          =   300
      Index           =   0
      Left            =   1950
      TabIndex        =   4
      Top             =   465
      Width           =   4185
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Porcentaje :"
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
      Index           =   4
      Left            =   60
      TabIndex        =   11
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta MN :"
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
      Top             =   510
      Width           =   870
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Columna :"
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
      Top             =   150
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   45
      X2              =   6345
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmMAsiDeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
Private nTipoAsi As Integer
Private pnCta_IndCCo As Integer
Private pcCodCCo_Def As String
Private sOrden As String
']

Private Sub cmdLlaveAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
    txtLlave(Index).SetFocus
  End Select
  ppAyuBus AYULLA, Index
End Sub

Private Sub Form_Load()
   Dim n_Index As Integer
   
   pbValidada = False

   Me.KeyPreview = True
   With frmMAsiGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain_1!codcta_mn.DefinedSize
      txtLlave(1).MaxLength = .uorstMain_1!codcta_me.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
      txtDato(0).MaxLength = .uorstMain_1!codcco.DefinedSize
      txtDato(1).MaxLength = 6
    ']
   End With
   nTipoAsi = frmMAsiGrd.uorstMain_0!tpoasi
   For n_Index = 0 To Choose(nTipoAsi, 10, 6, 4)
    If gsIdioma = NvlUsr_Sup Then
      If nTipoAsi = TPOGNR_CPR Then
        cboColumna.AddItem Choose(n_Index + 1, "Operación Gravada", "Operación Gravada/No Gravada", "Operación No Gravada", "Exonerado", "IGV", "ISC", "Otros", "Otros 1", "Otros 2", "Otros 3", "Total")
      ElseIf nTipoAsi = TPOGNR_VTA Then
        cboColumna.AddItem Choose(n_Index + 1, "Operación Gravada", "Exportación", "Exonerado", "IGV", "ISC", "Otros", "Total")
      ElseIf nTipoAsi = TPOGNR_HPR Then
        cboColumna.AddItem Choose(n_Index + 1, "Importe Bruto", "Retención 4ta Categoria", "Retención I.E.S.", "Otras Retenciones", "Importe Neto")
      End If
    Else
      If nTipoAsi = TPOGNR_CPR Then
        cboColumna.AddItem Choose(n_Index + 1, "Operation with Taxes", "Operation with/without Taxes", "Operation Without Taxes", "Discharged", "GST", "SCT", "Others", "Others 1", "Others 2", "Others 3", "Total")
      ElseIf nTipoAsi = TPOGNR_VTA Then
        cboColumna.AddItem Choose(n_Index + 1, "Operation with Taxes", "Export", "Discharged", "GST", "SCT", "Others", "Total")
      ElseIf nTipoAsi = TPOGNR_HPR Then
        cboColumna.AddItem Choose(n_Index + 1, "Gross Amount", "Withh.4th Class", "Withh. E.T.S.", "Others Withh.", "Net Amount")
      End If
    End If
   Next n_Index
   cboColumna.ListIndex = 0
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Columna :", "Cuenta MN :", "Cuenta ME :", "C.Costo :", "Porcentaje :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Column :", "Account NC :", "Account FC :", "C.Center :", "Percentage :")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']

End Sub

Private Sub Form_Activate()
  If txtLlave(0).Text <> "" Then
    ppAyuDet AYULLA, 0
    pnCta_IndCCo = frmMAsiGrd.uorstCoCta!IndCCo
    pcCodCCo_Def = IIf(IsNull(frmMAsiGrd.uorstCoCta!codcco_def), "", frmMAsiGrd.uorstCoCta!codcco_def)
    ' Actualiza los datos de centro de costo
    txtDato(0).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(1).Enabled)
    cmdDatoAyud(0).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(1).Enabled)
  End If
  If txtDato(0).Text <> "" Then ppAyuDet AYUDAT, 1
  
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMAsiGrd.uorstMain_1.BOF And frmMAsiGrd.uorstMain_1.EOF) Then
   frmMAsiGrd.uorstMain_1.CancelUpdate 'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMAsiGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMAsiGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
 
  '[Dato con el foco al corregir.       'Cambiar.
  txtDato(IIf(txtDato(0).Enabled, 0, 1)).SetFocus
  ']
End Sub

Public Sub cmdGrabar_Click()
  On Error GoTo Err
   
  Dim sSentencia As String
  Dim nRegistro As Long, nPorcentaje As Double
   
  '[Validacion de Datos segun Indicadores de Cuenta.
  If Len(Trim(txtLlave(0).Text)) = 0 Then
    MsgBox TEXT_6002, vbExclamation
    txtLlave(0).SetFocus
    Exit Sub
  End If
  If Len(Trim(txtLlave(1).Text)) = 0 Then
    MsgBox TEXT_6002, vbExclamation
    txtLlave(1).SetFocus
    Exit Sub
  End If
  ' Asignación de centro de costo
  If pnCta_IndCCo = INDCCO_ACT And Len(Trim(txtDato(0).Text)) = 0 Then
    MsgBox TEXT_6002, vbExclamation
    txtDato(0).SetFocus
    Exit Sub
  End If
  
  ' Verifico si existe cuenta con centro costo
  sSentencia = "SELECT COUNT(*) AS nRegistro "
  sSentencia = sSentencia & "FROM coasidet "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND codasi='" & frmMAsiGrd.uorstMain_0!codasi & "' "
  sSentencia = sSentencia & "AND tpocnc='" & (cboColumna.ListIndex + 1) & "' "
  sSentencia = sSentencia & "AND codcta_mn='" & txtLlave(0).Text & "' "
  sSentencia = sSentencia & "AND codcco='" & txtDato(0).Text & "' "
  sSentencia = sSentencia & "AND orden<>'" & sOrden & "'"
  nRegistro = CLng(gfRetornaValor(CONNSTRG & gsNomBDS, sSentencia))
  If nRegistro >= 1 Then
    MsgBox TEXT_8007, vbExclamation
    txtDato(0).SetFocus
    Exit Sub
  End If
  
  ' Verifico porcentaje distribución no sobrepase
  sSentencia = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(pordst), 0), 2) AS nPorcentaje "
  sSentencia = sSentencia & "FROM coasidet "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND codasi='" & frmMAsiGrd.uorstMain_0!codasi & "' "
  sSentencia = sSentencia & "AND tpocnc='" & (cboColumna.ListIndex + 1) & "' "
  sSentencia = sSentencia & "AND orden<>'" & sOrden & "'"
  nPorcentaje = CDec(gfRetornaValor(CONNSTRG & gsNomBDS, sSentencia))
  nPorcentaje = nPorcentaje + CDec(txtDato(1).Text)
  If nPorcentaje > 100 Then
    MsgBox Choose(gsIdioma, "porcentaje de distribución no valido", "Percentage of distribution not been worth"), vbExclamation
    txtDato(1).SetFocus
    Exit Sub
  End If
  
  With frmMAsiGrd                     'Cambiar Formulario de Grid.
    .uocnnMain.BeginTrans            'INICIA TRANSACCION.
    If pbNuevo Then
      .uorstMain_1.AddNew
    End If
    upDatosDesconectados 0
    
    With .uorstMain_1
      If pbNuevo Then
        !UsrCre = gsAbvUsr
        !FyHCre = Now
      Else
        !UsrMdf = gsAbvUsr
        !FyHMdf = Now
      End If
      .Update
    End With
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    If pbNuevo Then
      .uorstMain_1.Requery
      .upDatosGrid 1
      '[Búsqueda de llave actual.     'Cambiar.
      .uorstMain_1.Find "cLlave='" & Trim(cboColumna.ListIndex + 1) & txtLlave(0).Text & txtDato(0).Text & "'"
      ']
      cmdGrabar.Enabled = False
      upHabilitacion False
    
      upDatosPredeterminados
      '[Llave habilitar  'Cambiar.
      cboColumna.Enabled = True
      txtLlave(0).Enabled = True
      txtLlave(1).Enabled = True
      lblLlaveDeta(0).Enabled = True
      lblLlaveDeta(1).Enabled = True
      cmdLlaveAyud(0).Enabled = True
      cmdLlaveAyud(1).Enabled = True
      ']
      '[Llave con el foco al añadir.  'Cambiar.
      cboColumna.SetFocus
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
  
  frmMAsiGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
  
End Sub

Public Sub cmdDeshacer_Click()
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
  ppAyuBus AYUDAT, Index
End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    ppAyuBus AYULLA, Index
  End If
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
  If pbValidada Then                  'Cambiar.
    cboColumna.Enabled = False
    txtLlave(0).Enabled = False
    txtLlave(1).Enabled = False
    lblLlaveDeta(0).Enabled = False
    lblLlaveDeta(1).Enabled = False
    cmdLlaveAyud(0).Enabled = False
    cmdLlaveAyud(1).Enabled = False
    If txtDato(0).Enabled Then
      txtDato(0).SetFocus
    Else
      txtDato(1).SetFocus
    End If
  End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err

  Dim dvRegistro As Variant
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
  Select Case Index                   'Cambiar (añadir índices).
   Case 0, 1
    Cancel = ppAyuDet(AYULLA, Index)
    If Cancel Then Exit Sub
    If lblLlaveDeta(Index).Caption <> "" And Index = 0 Then
      pnCta_IndCCo = frmMAsiGrd.uorstCoCta!IndCCo
      pcCodCCo_Def = IIf(IsNull(frmMAsiGrd.uorstCoCta!codcco_def), "", frmMAsiGrd.uorstCoCta!codcco_def)
        
      ' Actualizo los datos adicionales
      txtDato(0).Text = IIf(txtDato(0).Text = "", pcCodCCo_Def, txtDato(0).Text)
      txtDato(0).Text = IIf(pnCta_IndCCo = INDCCO_ACT, txtDato(0).Text, "")
      lblDatoDeta(0).Caption = IIf(pnCta_IndCCo = INDCCO_ACT, lblDatoDeta(0).Caption, "")
    End If
  End Select
 
  'Valida la llave.                    'Cambiar.
  If Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0 And Index = 1 Then
    With frmMAsiGrd.uorstMain_1        'Cambiar Formulario de Grid.
      If Not (.BOF And .EOF) Then
        dvRegistro = .Bookmark
        .MoveFirst
        .Find "cLlave='" & Trim(cboColumna.ListIndex + 1) & txtLlave(0).Text & txtDato(0).Text & "'"
        If Not .EOF Then
          MsgBox TEXT_8007, vbExclamation
          If dvRegistro <> -1 Then .Bookmark = dvRegistro
          Cancel = True
          Exit Sub
        End If
        .Bookmark = dvRegistro
      End If
    End With
    
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
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    ppAyuBus AYUDAT, Index
  End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err

  'Asigna 0 a campos numéricos si están vacíos.
  If Index = 1 Then
    txtDato(Index).Text = IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)
    txtDato(Index).Text = IIf(CDec(txtDato(Index).Text) <= 0, 100, txtDato(Index).Text)
    txtDato(Index).Text = IIf(CDec(txtDato(Index).Text) >= 100, 100, txtDato(Index).Text)
    txtDato(Index).Text = FormatNumber(txtDato(Index).Text, 2)
  End If
  
  'Busca el dato en su tabla principal.
  If Index = 0 Then
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
  End If
      
  Exit Sub
Err:
  gpErrores
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYULLA Then
    Select Case tnIndex
     Case 0, 1                             'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
      txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
      lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  Else
    Select Case tnIndex
     Case 0                              'Cambiar (añadir índices).
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYULLA Then
    Select Case tnIndex                 'Cambiar.
     Case 0, 1
      If txtLlave(tnIndex).Text = "" Then
        lblLlaveDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmMAsiGrd.uorstCoCta
        .MoveFirst
        .Find "codcta='" & txtLlave(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblLlaveDeta(tnIndex).Caption = " " & frmMAsiGrd.uorstCoCta!detcta
        End If
      End With
    End Select
  Else
    Select Case tnIndex                 'Cambiar.
      Case 0
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmMAsiGrd.uorstCoCCo
        .MoveFirst
        .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & frmMAsiGrd.uorstCoCCo!DetCCo
        End If
      End With
    End Select
  End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  Dim porstOrden As ADODB.Recordset
   
  On Error GoTo Err
  
  With frmMAsiGrd                     'Cambiar Formulario de Grid.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        .uorstMain_1!codemp = frmMAsiGrd.uorstMain_0!codemp
        .uorstMain_1!pdoano = frmMAsiGrd.uorstMain_0!pdoano
        .uorstMain_1!codasi = frmMAsiGrd.uorstMain_0!codasi
        .uorstMain_1!tpocnc = (cboColumna.ListIndex + 1)
        .uorstMain_1!codcta_mn = txtLlave(0).Text
        .uorstMain_1!codcta_me = txtLlave(1).Text
        ' Obtengo el numero de orden
        Set porstOrden = New ADODB.Recordset
        With porstOrden
          .ActiveConnection = frmMAsiGrd.uocnnMain
          .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(orden), '00') AS cOrden "
          .Source = .Source & "FROM coasidet "
          .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
          .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
          .Source = .Source & "AND codasi='" & frmMAsiGrd.uorstMain_0!codasi & "' "
          .Source = .Source & "AND tpocnc='" & (cboColumna.ListIndex + 1) & "'"
          '     .CursorLocation = adUseClient   'Es el Default.
          .CursorType = adOpenForwardOnly
          .LockType = adLockReadOnly
          .Open
          frmMAsiGrd.uorstMain_1!orden = Format(Val(porstOrden!cOrden) + 1, "00")
          .Close
        End With
      End If

      'Datos.
      .uorstMain_1!codcco = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      .uorstMain_1!pordst = CDec(txtDato(1).Text)
    Else
      'Llaves.
      cboColumna.ListIndex = (.uorstMain_1!tpocnc - 1)
      txtLlave(0).Text = .uorstMain_1!codcta_mn
      txtLlave(1).Text = .uorstMain_1!codcta_me
      sOrden = .uorstMain_1!orden
      
      'Datos.
      txtDato(0).Text = IIf(IsNull(.uorstMain_1!codcco), "", .uorstMain_1!codcco)
      txtDato(1).Text = Format(IIf(IsNull(.uorstMain_1!pordst), 0, .uorstMain_1!pordst), FORMATO_NUM_2)
         
      'Busca detalle de códigos.
      ppAyuDet AYULLA, 0
      ppAyuDet AYULLA, 1
      ppAyuDet AYUDAT, 0
    End If
  End With
  Set porstOrden = Nothing
      
  Exit Sub
Err:
  gpErrores
   
  Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
   Dim dnContador As Integer

  'Llaves.
   cboColumna.ListIndex = 0
   txtLlave(0).Text = ""
   txtLlave(1).Text = ""
   sOrden = "00"
   pbValidada = False

  'Datos.
   txtDato(0).Text = ""
   txtDato(1).Text = Format(0, FORMATO_NUM_2)

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
   lblLlaveDeta(1).Caption = ""
   lblDatoDeta(0).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  txtDato(0).Enabled = (tbHabilitar And pnCta_IndCCo = INDCCO_ACT)
  
  'Ayudas.
  cmdLlaveAyud(0).Enabled = (pbNuevo)
  cmdLlaveAyud(1).Enabled = (pbNuevo)
  lblLlaveDeta(0).Enabled = tbHabilitar
  lblLlaveDeta(1).Enabled = tbHabilitar
  cmdDatoAyud(0).Enabled = (tbHabilitar And pnCta_IndCCo = INDCCO_ACT)
  lblDatoDeta(0).Enabled = tbHabilitar

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


