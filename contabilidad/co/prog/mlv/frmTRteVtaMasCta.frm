VERSION 5.00
Begin VB.Form frmTRteVtaMasCta 
   Caption         =   "[Entidad]"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   7620
   StartUpPosition =   1  'CenterOwner
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
      Height          =   870
      Index           =   2
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   840
      Width           =   6525
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
      Height          =   870
      Index           =   3
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1755
      Width           =   6525
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   7290
      Picture         =   "frmTRteVtaMasCta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   480
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
      Index           =   1
      Left            =   960
      MaxLength       =   11
      TabIndex        =   4
      Top             =   480
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2130
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3060
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
         Picture         =   "frmTRteVtaMasCta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmTRteVtaMasCta.frx":02F4
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmTRteVtaMasCta.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmTRteVtaMasCta.frx":04F8
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmTRteVtaMasCta.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmTRteVtaMasCta.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##,##0.00"
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
      Index           =   4
      Left            =   960
      TabIndex        =   11
      Top             =   2670
      Width           =   1815
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7290
      Picture         =   "frmTRteVtaMasCta.frx":0996
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   135
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
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   975
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   1815
      Width           =   900
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
      Index           =   1
      Left            =   2235
      TabIndex        =   5
      Top             =   480
      Width           =   5070
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Auxiliar :"
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
      TabIndex        =   3
      Top             =   480
      Width           =   630
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   900
      Width           =   510
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
      TabIndex        =   10
      Top             =   2730
      Width           =   795
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   600
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
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   5385
   End
End
Attribute VB_Name = "frmTRteVtaMasCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
Private ps_OrdCuenta As String          ' Orden de cuenta
']

Private Sub Form_Load()
  pbValidada = False
  
  Me.KeyPreview = True
  
  With frmTRteVtaGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
    txtDato(0).MaxLength = .uorstCoRteVtaCta!codcta.DefinedSize
    ']
    '[Datos                            'Cambiar.
    txtDato(1).MaxLength = 11
    txtDato(2).MaxLength = 250
    txtDato(3).MaxLength = 250
    txtDato(4).MaxLength = 14
  End With
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  cmdCorregir.Enabled = (Not pbNuevo)
  cmdRetroceder.Enabled = (Not pbNuevo)
  cmdAvanzar.Enabled = (Not pbNuevo)
  upHabilitacion pbNuevo
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuenta :", "Auxiliar :", "Glosa :", "Traducción :", "Porcentaje :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Account :", "Auxiliary :", "Gloss  :", "Translation :", "Percentage :")
  Next nElemento
  cmdGrabar.Caption = Choose(gsIdioma, "&Aceptar", "&Accept")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, False, True, True, aLabel
  ']
End Sub

Private Sub Form_Activate()
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
''   Call gpTeclasData2(KeyAscii)
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
  upHabilitacion (True)
  
  '[Dato con el foco al corregir.       'Cambiar.
  txtDato(1).SetFocus
  ']
End Sub

Private Sub cmdGrabar_Click()
  On Error GoTo Err
  
  With frmTRteVtaGrd                     'Cambiar Formulario de Grid.
    '      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
    If pbNuevo Then
      .uorstCoRteVtaCta.AddNew
    End If
    upDatosDesconectados 0
    With .uorstCoRteVtaCta
      If pbNuevo Then
        !UsrCre = gsAbvUsr
        !FyHCre = Now
      Else
        !UsrMdf = gsAbvUsr
        !FyHMdf = Now
      End If
      .Update
    End With
  
    If pbNuevo Then
      .uorstCoRteVtaCta.Requery
      .upDatosGrid
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      cmdAvanzar.Enabled = False
      cmdRetroceder.Enabled = False
      cmdCorregir.Enabled = False
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
  
  '   frmTrteVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
   Case 1
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    ppAyuBus Index
  End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  Dim dvRegistroActual As Variant
  
  Select Case Index
   Case 4
    If txtDato(Index).Text = "" Then
      txtDato(Index).Text = 0
    End If
  End Select
  
  'Da formato.
  Select Case Index
   Case 4
    txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
  End Select
  
  'Busca el dato en su tabla principal.
  Select Case Index
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      '[
      With frmTRteVtaGrd.uorstCoRteVtaCta
        If Not (.BOF Or .EOF) And .RecordCount > 0 Then
          dvRegistroActual = .Bookmark
          .MoveFirst
          .Find "cLlave2='" & frmTRteVta.txtLlave(0).Text & frmTRteVta.txtLlave(1).Text & frmTRteVtaMasGrd.unIndice & ps_OrdCuenta & txtDato(0).Text & "'"
          If Not .EOF Then
            MsgBox TEXT_8007, vbExclamation
            If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
            Cancel = True
            Exit Sub
          End If
          .Bookmark = dvRegistroActual
        End If
      End With
      cmdGrabar.Enabled = True
      upHabilitacion True
    Else
      upHabilitacion False
      cmdGrabar.Enabled = False
    End If
    cmdDatoAyud(0).Enabled = True
    ']
   Case 1
    If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      '[
      With frmTRteVtaGrd.uorstCoRteVtaCta
        If Not (.BOF Or .EOF) And .RecordCount > 0 Then
          dvRegistroActual = .Bookmark
          .MoveFirst
          .Find "cLlave2='" & frmTRteVta.txtLlave(0).Text & frmTRteVta.txtLlave(1).Text & frmTRteVtaMasGrd.unIndice & ps_OrdCuenta & txtDato(1).Text & "'"
          If Not .EOF Then
            MsgBox TEXT_8007, vbExclamation
            If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
            Cancel = True
            Exit Sub
          End If
          .Bookmark = dvRegistroActual
        End If
      End With
      cmdGrabar.Enabled = True
      upHabilitacion True
    Else
      upHabilitacion True
      cmdGrabar.Enabled = True
    End If
    cmdDatoAyud(0).Enabled = True
    ']
  End Select
  
  Exit Sub
Err:
  gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
    modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1                           'Cambiar (añadir índices).
    modAyuBus.Aux_Det "IndCli=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
    With frmTRteVtaGrd.uorstCOCta
      .MoveFirst
      .Find "CodCta='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
      End If
    End With
   Case 1
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmTRteVtaGrd.uorstTGAux
      If .RecordCount > 0 Then .MoveFirst
        If Len(Trim(txtDato(tnIndex).Text)) <> 0 Then
          .Find "CodAux='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & !razAux
        End If
      End If
    End With
  End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  On Error GoTo Err
  
  With frmTRteVtaGrd.uorstCoRteVtaCta    'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !sernegocio = frmTRteVta.txtLlave(0).Text
        !nronegocio = frmTRteVta.txtLlave(1).Text
        !tpocnc = frmTRteVtaMasGrd.unIndice
        !orden = ps_OrdCuenta
      End If
      
      'Datos.
      !codcta = IIf(txtDato(0).Text = "", Null, txtDato(0))
      !codruc = IIf(txtDato(1).Text = "", Null, txtDato(1))
      !glodet0 = Left(IIf(txtDato(Choose(gsIdioma, 2, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 2, 3)).Text), 250)
      !glodet1 = Mid(IIf(txtDato(Choose(gsIdioma, 2, 3)).Text = "" Or Len(txtDato(Choose(gsIdioma, 2, 3)).Text) <= 250, Null, txtDato(Choose(gsIdioma, 2, 3)).Text), 251)
      !glodet0x = Left(IIf(txtDato(Choose(gsIdioma, 3, 2)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 2)).Text), 250)
      !glodet1x = Mid(IIf(txtDato(Choose(gsIdioma, 3, 2)).Text = "" Or Len(txtDato(Choose(gsIdioma, 3, 2)).Text) <= 250, Null, txtDato(Choose(gsIdioma, 3, 2)).Text), 251)
      !porimpcta = CDec(txtDato(4).Text)
    Else
      'Datos.
      txtDato(0).Text = IIf(IsNull(!codcta), "", !codcta)
      txtDato(1).Text = IIf(IsNull(!codruc), "", !codruc)
      txtDato(Choose(gsIdioma, 2, 3)).Text = IIf(IsNull(!glodet), "", !glodet)
      txtDato(Choose(gsIdioma, 3, 2)).Text = IIf(IsNull(!glodetx), "", !glodetx)
      txtDato(4).Text = Format(!porimpcta, FORMATO_NUM_1)
    End If
  End With
  Exit Sub
Err:
  gpErrores
  
  Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
  Dim dnContador As Integer
  
  'Datos.
  txtDato(0).Text = ""
  txtDato(1).Text = ""
  txtDato(2).Text = ""
  txtDato(3).Text = ""
  txtDato(4).Text = Format(0, FORMATO_NUM_1)
  ps_OrdCuenta = "00"
  
  If pbNuevo Then
    txtDato(2).Text = frmTRteVta.txtDato(2).Text
    txtDato(4).Text = Format(100, FORMATO_NUM_1)
    With frmTRteVtaGrd.uorstCoRteVtaCta
      If Not .EOF And .RecordCount > 0 Then
        .MoveFirst
        Do
          txtDato(4).Text = Format(CDec(txtDato(4).Text) - !porimpcta, FORMATO_NUM_1)
          ps_OrdCuenta = IIf(!orden > ps_OrdCuenta, !orden, ps_OrdCuenta)
          .MoveNext
        Loop Until .EOF
        .MoveFirst
      End If
    End With
  End If
  ps_OrdCuenta = gfCeros(ps_OrdCuenta, 2, 1, "0")
  'Ayudas.
  lblDatoDeta(0).Caption = ""
  lblDatoDeta(1).Caption = ""

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
    Next
  End With
  'Ayudas.
  cmdDatoAyud(0).Enabled = tbHabilitar
  lblDatoDeta(0).Enabled = tbHabilitar
  cmdDatoAyud(1).Enabled = tbHabilitar
  lblDatoDeta(1).Enabled = tbHabilitar

End Sub

'[Propio del formulario.
']

Public Property Get zbNuevo() As Boolean
  zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
  pbNuevo = tbNuevo
End Property
