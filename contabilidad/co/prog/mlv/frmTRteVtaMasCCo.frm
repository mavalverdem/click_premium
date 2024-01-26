VERSION 5.00
Begin VB.Form frmTRteVtaMasCCo 
   Caption         =   "[Entidad]"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
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
      Index           =   1
      Left            =   1005
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   922
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1020
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
         Picture         =   "frmTRteVtaMasCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmTRteVtaMasCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
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
         Picture         =   "frmTRteVtaMasCCo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmTRteVtaMasCCo.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTRteVtaMasCCo.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmTRteVtaMasCCo.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   5040
      Picture         =   "frmTRteVtaMasCCo.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   615
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
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "C.Costo:"
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
      TabIndex        =   10
      Top             =   180
      Width           =   615
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
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   540
      Width           =   855
   End
End
Attribute VB_Name = "frmTRteVtaMasCCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
']

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmTRteVtaGrd
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstCoRteVtaCCo!codcco.DefinedSize
      txtDato(1).MaxLength = 14
    ']
   End With
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   cmdRetroceder.Enabled = (Not pbNuevo)
   cmdCorregir.Enabled = (Not pbNuevo)
   cmdAvanzar.Enabled = (Not pbNuevo)
   upHabilitacion pbNuevo

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "C.Costo :", "Porcentaje :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "C.Center :", "Percentage :")
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
  
  With frmTRteVtaGrd                      'Cambiar Formulario de Grid.
    If pbNuevo Then
      .uorstCoRteVtaCCo.AddNew
    End If
    upDatosDesconectados 0
    With .uorstCoRteVtaCCo
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
      .uorstCoRteVtaCCo.Requery
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
  If KeyCode = vbKeyF2 Then
    ppAyuBus Index
  End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  Dim dnContador As Integer
  Dim dvRegistroActual As Variant
  
  On Error GoTo Err
  Select Case Index
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
    End If
  End Select
  
  Select Case Index
    Case 1
    If txtDato.Item(Index).Text = "" Then
      txtDato.Item(Index).Text = 0
    End If
  End Select

  'Da formato.
   Select Case Index
   Case 1, 2
      txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
  Select Case Index
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      With frmTRteVtaGrd.uorstCoRteVtaCCo
        If Not (.BOF Or .EOF) And .RecordCount > 0 Then
          dvRegistroActual = .Bookmark
          .MoveFirst
          .Find "cLlave2='" & frmTRteVta.txtLlave(0).Text & frmTRteVta.txtLlave(1).Text & frmTRteVtaMasGrd.unIndice & frmTRteVtaGrd.uorstCoRteVtaCta!orden & frmTRteVtaMasGrd.dgrMain.Columns(0).Text & txtDato(0).Text & "'"
          If Not .EOF Then
            MsgBox TEXT_8007, vbExclamation
            If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
            Cancel = True
            Exit Sub
          End If
          .Bookmark = dvRegistroActual
        End If
      End With
      upHabilitacion True
      cmdGrabar.Enabled = True
    Else
      upHabilitacion False
      cmdGrabar.Enabled = False
    End If
    cmdDatoAyud(0).Enabled = True
  End Select
      
  Exit Sub
Err:
  gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
    modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
    With frmTRteVtaGrd.uorstCoCCo
      .MoveFirst
      .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
      End If
    End With
  End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
  On Error GoTo Err
  
  With frmTRteVtaGrd.uorstCoRteVtaCCo    'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !sernegocio = frmTRteVta.txtLlave(0).Text
        !nronegocio = frmTRteVta.txtLlave(1).Text
        !tpocnc = frmTRteVtaMasGrd.unIndice
        !orden = frmTRteVtaGrd.uorstCoRteVtaCta!orden
        !codcta = frmTRteVtaMasGrd.dgrMain.Columns(0).Text
      End If
      !codcco = IIf(txtDato(0).Text = "", Null, txtDato(0))
      !porimpcco = CDec(txtDato(1).Text)
    Else
      txtDato(0).Text = IIf(IsNull(!codcco), "", !codcco)
      txtDato(1).Text = Format(!porimpcco, FORMATO_NUM_1)
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
  txtDato(1).Text = Format(0, FORMATO_NUM_1)
  If pbNuevo Then
    txtDato(1).Text = Format(100, FORMATO_NUM_1)
    With frmTRteVtaGrd.uorstCoRteVtaCCo
      If Not .EOF And .RecordCount > 0 Then
        .MoveFirst
        Do
          txtDato(1).Text = Format(CDec(txtDato(1).Text) - !porimpcco, FORMATO_NUM_1)
          .MoveNext
        Loop Until .EOF
        .MoveFirst
      End If
    End With
  End If
  lblDatoDeta(dnContador).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Integer

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
End Sub
'[Propio del formulario.
']
Public Property Get zbNuevo() As Boolean
  zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
  pbNuevo = tbNuevo
  'Habilita o no los botones.
  cmdDeshacer.Enabled = IIf(pbNuevo = True, False, True)
End Property
