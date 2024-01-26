VERSION 5.00
Begin VB.Form frmMProd 
   Caption         =   "[Entidad]"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   3
      Left            =   7200
      Picture         =   "frmMProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2265
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
      Index           =   2
      Left            =   1110
      TabIndex        =   7
      Top             =   1920
      Width           =   1035
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
      Height          =   540
      Index           =   1
      Left            =   1110
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1335
      Width           =   6390
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   60
      TabIndex        =   16
      Top             =   3360
      Width           =   1275
      Begin VB.CheckBox chkEstProd 
         Caption         =   "Activo"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Frame fraCuadro 
      Caption         =   " Precio Compra "
      ForeColor       =   &H00800000&
      Height          =   660
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   2685
      Width           =   7410
      Begin VB.ComboBox cboTpoMon 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmMProd.frx":01AA
         Left            =   1380
         List            =   "frmMProd.frx":01AC
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   225
         Width           =   1980
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   4815
         TabIndex        =   15
         Top             =   225
         Width           =   2190
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
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
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   660
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Importe  :"
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
         Left            =   3885
         TabIndex        =   14
         Top             =   285
         Width           =   660
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
      Index           =   3
      Left            =   1110
      TabIndex        =   9
      Top             =   2265
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2010
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3615
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
         Picture         =   "frmMProd.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "frmMProd.frx":0358
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Picture         =   "frmMProd.frx":0502
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmMProd.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmMProd.frx":074E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdSalir 
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
         Picture         =   "frmMProd.frx":0850
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
      Height          =   540
      Index           =   0
      Left            =   1110
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   750
      Width           =   6390
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
      Left            =   825
      TabIndex        =   1
      Top             =   120
      Width           =   2145
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
      Height          =   315
      Index           =   3
      Left            =   2115
      TabIndex        =   10
      Top             =   2265
      Width           =   5130
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Unid. Medida :"
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
      Left            =   75
      TabIndex        =   6
      Top             =   1995
      Width           =   1005
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
      Left            =   75
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta :"
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
      Left            =   60
      TabIndex        =   8
      Top             =   2325
      Width           =   600
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   945
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
      Top             =   180
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   7440
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

Private Sub Form_Load()
  pbValidada = False
  Dim n_Contador As Integer
  
  Me.KeyPreview = True
  
  With frmMProdGrd
    '[Llaves                           'Cambiar
    txtLlave(0).MaxLength = .uorstMain!codprod.DefinedSize
    ']
    '[Datos                            'Cambiar.
    txtDato(0).MaxLength = .uorstMain!detprod.DefinedSize
    txtDato(1).MaxLength = .uorstMain!detprodx.DefinedSize
    txtDato(2).MaxLength = .uorstMain!unimed.DefinedSize
    txtDato(3).MaxLength = .uorstMain!codcta.DefinedSize
    txtDato(4).MaxLength = 17
    ']
  End With
  ' Configuro tipo moneda
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND
   
  If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
  End If
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(7, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Producto :", "Descripción :", "Traducción :", "Unid Medida :", "Cuenta :", "Moneda :", "Importe :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Product :", "Description :", "Traslation :", "Measure Unit :", "Account :", "Currency :", "Amount :")
  Next nElemento
  fraCuadro(0).Caption = Choose(gsIdioma, " Precio Compra ", " Purchase Price ")
  chkEstProd.Caption = Choose(gsIdioma, "&Activo", "&Active")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']

End Sub

Private Sub Form_Activate()
  '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMProdGrd.uorstMain.BOF And frmMProdGrd.uorstMain.EOF) Then
   frmMProdGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMProdGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMProdGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
  Dim nOpbIndex As Integer
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
 '[Dato con el foco al corregir.       'Cambiar.
   If txtDato(0).Enabled Then
      txtDato(0).SetFocus
   Else
      txtDato(1).SetFocus
   End If
 ']
End Sub

Public Sub cmdGrabar_Click()
  
  On Error GoTo Err
  
  If Len(Trim(txtLlave(0).Text)) = 0 Then MsgBox TEXT_8005, vbExclamation: txtLlave(0).SetFocus: Exit Sub
  If Len(Trim(txtDato(0).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If Len(Trim(txtDato(2).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(2).SetFocus: Exit Sub
  If Len(Trim(txtDato(3).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(3).SetFocus: Exit Sub
  If CDec(txtDato(4).Text) <= 0 Then MsgBox TEXT_8010, vbExclamation: txtDato(4).SetFocus: Exit Sub
  
  With frmMProdGrd                     'Cambiar Formulario de Grid.
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
        !FyHMdf = Now
      End If
      .Update
    End With
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  
    If pbNuevo Then
      .uorstMain.Requery
      .ppDatosGrid
      '[Búsqueda de llave actual.     'Cambiar.
      .uorstMain.Find "codprod='" & txtLlave(0).Text & "'"
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
  frmMProdGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.

End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)

  Select Case Index                   'Cambiar. Añadir índices.
   Case 3
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub
Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
  'Cambiar.
  If pbValidada Then txtDato(0).SetFocus
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  Dim dvRegistro As Variant
  Dim nOpbIndex As Integer

'   On Error GoTo Err
  'Valida la llave.                    'Cambiar.
  If Len(Trim(txtLlave(Index).Text)) <> 0 Then
    With frmMProdGrd.uorstMain
      If Not (.BOF And .EOF) Then
        dvRegistro = .Bookmark
        .MoveFirst
        .Find "codprod='" & txtLlave(0).Text & "'"
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
'Err:
'   gpErrores
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
  If KeyCode = vbKeyF2 Then
    ppAyuBus Index
  End If
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  Select Case Index
   Case 3           ' busca dato en tabla principal
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   Case 4           ' ceros a la izquierda
    txtDato(Index).Text = IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)
    txtDato(Index).Text = FormatNumber(txtDato(Index).Text, 2)
  End Select

  Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 3                              'Cambiar (añadir índices).
    modAyuBus.Cta_Cod "tpocta=" & TPOCTA_TRA & " AND estcta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
  
  Select Case tnIndex                 'Cambiar.
   Case 3
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmMProdGrd.uorstCoCta
      .MoveFirst
      .Find "codcta='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
      End If
    End With
  End Select

End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  
  On Error GoTo Err

  With frmMProdGrd
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        .uorstMain!codemp = gsCodEmp
        .uorstMain!pdoano = gsAnoAct
        .uorstMain!codprod = txtLlave(0).Text
      End If
      'Datos.
      .uorstMain!detprod = IIf(txtDato(gsIdioma - 1).Text = "", Null, txtDato(gsIdioma - 1).Text)
      .uorstMain!detprodx = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
      .uorstMain!unimed = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      .uorstMain!codcta = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
      .uorstMain!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      .uorstMain!impcpr = CDec(txtDato(4).Text)
      .uorstMain!estprod = IIf(chkEstProd.Value = vbChecked, ESTCTA_ACT, ESTCTA_INA)
    Else
      'Llaves.
      txtLlave(0).Text = .uorstMain!codprod
      'Datos.
      txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain!detprod), "", .uorstMain!detprod)
      txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain!detprodx), "", .uorstMain!detprodx)
      txtDato(2).Text = IIf(IsNull(.uorstMain!unimed), "", .uorstMain!unimed)
      txtDato(3).Text = IIf(IsNull(.uorstMain!codcta), "", .uorstMain!codcta)
      cboTpoMon.ListIndex = IIf(.uorstMain!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      txtDato(4).Text = Format(IIf(IsNull(.uorstMain!impcpr), 0, .uorstMain!impcpr), FORMATO_NUM_1)
      chkEstProd.Value = IIf(.uorstMain!estprod = ESTCTA_ACT, vbChecked, vbUnchecked)
      ppAyuDet 3
    End If
  End With
      
  Exit Sub
Err:
  gpErrores
 '  Resume

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
    .Item(dnContador - 1).Text = Format(0, FORMATO_NUM_1)
  End With
  chkEstProd.Value = vbChecked
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  'Ayudas.
  lblDatoDeta(3).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Integer

  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  cboTpoMon.Enabled = tbHabilitar
  chkEstProd.Enabled = tbHabilitar
  cmdDatoAyud(3).Enabled = tbHabilitar

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
