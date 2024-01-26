VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTCon 
   Caption         =   "[Título]"
   ClientHeight    =   4005
   ClientLeft      =   900
   ClientTop       =   1650
   ClientWidth     =   8385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   8385
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2355
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3315
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
         Picture         =   "frmTCon.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Picture         =   "frmTCon.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Picture         =   "frmTCon.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "frmTCon.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "frmTCon.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frmTCon.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   5
      Left            =   3165
      TabIndex        =   18
      Top             =   2865
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   4
      Left            =   1080
      TabIndex        =   16
      Top             =   2865
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   3
      Left            =   2880
      TabIndex        =   14
      Top             =   2205
      Width           =   735
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   2
      Left            =   1080
      TabIndex        =   10
      Top             =   1830
      Width           =   7050
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   1470
      Width           =   7050
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   1280
   End
   Begin VB.TextBox txtLlave 
      Height          =   280
      Index           =   0
      Left            =   1470
      TabIndex        =   1
      Top             =   210
      Width           =   1140
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "Cliente"
      Height          =   315
      Left            =   7035
      TabIndex        =   26
      Top             =   1110
      Width           =   1110
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   7880
      Picture         =   "frmTCon.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTCon.frx":0996
      Left            =   1080
      List            =   "frmTCon.frx":0998
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2205
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   99024897
      CurrentDate     =   37102
   End
   Begin VB.Shape shpCuadro 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   555
      Left            =   60
      Top             =   75
      Width           =   8235
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe ME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   8
      Left            =   3195
      TabIndex        =   17
      Top             =   2625
      Width           =   780
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe MN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   7
      Left            =   1095
      TabIndex        =   15
      Top             =   2625
      Width           =   795
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
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   4
      Left            =   60
      TabIndex        =   9
      Top             =   1875
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   1125
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8300
      Y1              =   630
      Y2              =   630
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
      Left            =   2340
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Glosa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   1515
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   5
      Left            =   60
      TabIndex        =   11
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "T.Cambio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   2235
      Width           =   705
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Orden Servicio :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   1170
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   765
      Width           =   570
   End
End
Attribute VB_Name = "frmTCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbCorregir As Boolean
Private pbValidada As Boolean
Private pbFecha As Boolean

Private pnCta_TpoTcb As String

Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2

']
Private Sub Form_Load()
  pbValidada = False
  pbFecha = True
  Me.KeyPreview = True
  
  With frmTConGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
    txtLlave(0).MaxLength = .uorstMain!codcon.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
    End With
    
    txtDato(0).MaxLength = .uorstMain!codaux.DefinedSize
    txtDato(gsIdioma).MaxLength = .uorstMain!detcon.DefinedSize
    txtDato(3 - gsIdioma).MaxLength = .uorstMain!detconx.DefinedSize
    txtDato(3).MaxLength = 7
    txtDato(4).MaxLength = 14
    txtDato(5).MaxLength = 14
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
  ReDim aLabel(8, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Orden Servicio :", "Cliente :", "Fecha :", "Glosa :", "Traducción :", "Moneda :", "T.Cambio:", "Importe MN :", "Importe ME :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Service Order :", "Customer :", "Date :", "Gloss :", "Translation :", "Currency :", "R.Exchange :", "Amount NC :", "Amount FC :")
  Next nElemento
  cmdAuxiliar.Caption = Choose(gsIdioma, "Cliente", "Customer")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']

  '[Propio del formulario.
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  ']
End Sub

Private Sub Form_Activate()
  
 '[Busca detalle de códigos.           'Cambiar (habilitar/deshabilitar).
  If txtDato(0).Text <> "" Then ppAyuDet AYUDAT, 0
 ']
  
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If

 '[Propio del formulario.
  If Not pbNuevo Then
    dtpDato.Tag = dtpDato.Value
  End If
  txtDato(3).Tag = txtDato(3).Text
 ']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not frmTConGrd.uorstMain.EOF Then
    If frmTConGrd.uorstMain.EditMode <> adEditNone Then frmTConGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmTConGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTConGrd.uorstMain_Grd.MoveFirst
   frmTConGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtDato(0).Text & "'"
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmTConGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTConGrd.uorstMain_Grd.MoveFirst
   frmTConGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtDato(0).Text & "'"
End Sub

Public Sub cmdCorregir_Click()
  'Verificación de Mes Cerrado.
  If (gbCieVta Or gbCieCpb Or gbCieCpr Or gbCieHpr) Then MsgBox TEXT_9016, vbCritical: Exit Sub
  
  pbCorregir = True
  
  cmdRetroceder.Enabled = False
  cmdAvanzar.Enabled = False
  cmdCorregir.Enabled = False
  cmdGrabar.Enabled = True
  cmdDeshacer.Enabled = True
  upHabilitacion True
  txtDato(0).Enabled = False
  cmdDatoAyud(0).Enabled = False
  
  '[Dato con el foco al corregir.       'Cambiar.
  dtpDato.SetFocus
  ']
  
End Sub

Public Sub cmdGrabar_Click()
  Dim sSentencia As String
  On Error GoTo Err
  
  If Len(Trim(txtDato(0).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If cboTpoMon.ListIndex = TPOMON_NAC_IND And CDec(txtDato(4).Text) = 0 Then
    MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Nacional.", "You Must enter the amount in National Currency."), vbInformation
    txtDato(4).SetFocus
    Exit Sub
  ElseIf cboTpoMon.ListIndex = TPOMON_EXT_IND And CDec(txtDato(5).Text) = 0 Then
    MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Extranjera.", "You Must enter the amount in Foreign Currency."), vbInformation
    txtDato(5).SetFocus
    Exit Sub
  End If
   
  With frmTConGrd                     'Cambiar Formulario de Grid.
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
    
    ' Refresco la grilla verificar
    .uorstMain_Grd.Requery
    .upDatosGrid
    '[Búsqueda de llave actual.     'Cambiar.
    .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtDato(0).Text & "'"
    ']
    If pbNuevo Then
      pbValidada = False
      cmdGrabar.Enabled = False
      upHabilitacion False
      txtLlave(0).Enabled = False
      '[ No Pertenece al Formulario
      .uorstMain.Requery
      
      upDatosPredeterminados
      txtLlave(0).Enabled = True
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
  
   frmTConGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Public Sub cmdSalir_Click()
  If pbNuevo Or pbCorregir Then pbCorregir = False
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
'''[ARREGLAR: Retrocede si Shift está presionado.
''   If Len(Trim(txtLlave(Index))) + 1 = txtLlave(Index).MaxLength Then
''      SendKeys "{TAB}"
''   End If
''']ARREGLAR.
 
 '[Convierte a mayúsculas.
'   If Index = 0 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub
Private Sub txtLlave_LostFocus(Index As Integer)
  If pbValidada Then
    txtLlave(0).Enabled = False
    If txtDato(0).Enabled Then
      txtDato(0).SetFocus
    ElseIf dtpDato.Enabled Then
      dtpDato.SetFocus
    End If
  End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  Dim dvRegistro As Variant
  Dim sSentencia As String
  
  ' Valido la llave
  If Len(Trim(txtLlave(0).Text)) <> 0 Then
    With frmTConGrd                  'Cambiar Formulario de Grid.
      sSentencia = "SELECT mespvs FROM coconser "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND codcon='" & txtLlave(0).Text & "'"
      Set .porstCancel = .uocnnMain.Execute(sSentencia)
      If .porstCancel.RecordCount > 0 Then
        MsgBox TEXT_8007 & Chr(13) & Choose(gsIdioma, "(mes ", "(month ") & gfMesLet("01" & .porstCancel!mespvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
        Cancel = True
        Exit Sub
      End If
      .porstCancel.Close
    End With
    
    With frmTConGrd.uorstMain
      If Not (.BOF And .EOF) Then
        dvRegistro = .Bookmark
        .MoveFirst
        .Find "cLlave='" & txtLlave(0).Text & "'"
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
  If Index >= 4 And Index <= 5 Then
    If Val(txtDato(3).Text) = 0 Then
      txtDato(3).Text = Format(0, FORMATO_NUM_2)
      txtDato(3).SetFocus
      MsgBox TEXT_9015, vbExclamation
      Exit Sub
    End If
  End If
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength + IIf(Index >= 3 And Index <= 6, 1, 0)
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  '[ARREGLAR: Retrocede si Shift está presionado.
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus AYUDAT, Index
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
   
  Select Case Index
   Case 3
    If Val(txtDato(Index).Text) > 0 Then
      txtDato(Index).Text = Format(Val(txtDato(Index).Text), FORMATO_NUM_2)
    End If
   Case 4, 5
    If CDec(txtDato(3).Text) <= 0 Then
      MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
      txtDato(3).SetFocus
      Exit Sub
    End If
    ' Convierto importe en cero
    If CDec(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      If Index = 4 And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index + 1).Text) * CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 5 And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index - 1).Text) / CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    ElseIf CDec(txtDato(Index).Text) <> 0 Then
      If Index = 4 And cboTpoMon.ListIndex = TPOMON_NAC_IND And (txtDato(Index - 1).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index + 1).Text = Format(Round(CDec(txtDato(Index).Text) / CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 5 And cboTpoMon.ListIndex = TPOMON_EXT_IND And (txtDato(Index - 1).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index - 1).Text = Format(Round(CDec(txtDato(Index).Text) * CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    End If
   End Select

End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  'Completa con ceros a la izquierda.
  Select Case Index
   Case 0
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
   Case 3
      txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_2)
   Case 4, 5
      txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_1)
  End Select
   
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYUDAT Then
    Select Case tnIndex
     Case 0                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "IndCli=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYUDAT Then
    Select Case tnIndex                 'Cambiar.
     Case 0
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTConGrd.uorstTGAux
        If .RecordCount > 0 Then .MoveFirst
        .Find "codaux='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!RazAux), "", !RazAux)
        End If
      End With
    End Select
  End If
  
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  
  On Error GoTo Err

  '[Propio del formulario.
  Dim dnContador As Byte
  ']
  With frmTConGrd.uorstMain           'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !mespvs = gsMesAct
        !codcon = txtLlave(0).Text
      End If

      'Datos.
      !codaux = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !fehcon = dtpDato.Value
      !detcon = IIf(txtDato(gsIdioma).Text = "", Null, txtDato(gsIdioma).Text)
      !detconx = IIf(txtDato(3 - gsIdioma).Text = "", Null, txtDato(3 - gsIdioma).Text)
      !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
      !ImpTCb = CDec(txtDato(3).Text)
      !ImpMN = CDec(txtDato(4).Text)
      !ImpME = CDec(txtDato(5).Text)
    Else
      'Llaves.
      txtLlave(0).Text = !codcon
      
      'Datos.
      txtDato(0).Text = IIf(IsNull(!codaux), "", !codaux)
      dtpDato.Value = !fehcon
      txtDato(gsIdioma).Text = IIf(IsNull(!detcon), "", !detcon)
      txtDato(3 - gsIdioma).Text = IIf(IsNull(!detconx), "", !detconx)
      cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      txtDato(3).Text = Format(!ImpTCb, FORMATO_NUM_2)
      txtDato(4).Text = Format(!ImpMN, FORMATO_NUM_1)
      txtDato(5).Text = Format(!ImpME, FORMATO_NUM_1)
      
      txtDato(4).Tag = Format(txtDato(4).Text, FORMATO_NUM_1)
      txtDato(5).Tag = Format(txtDato(5).Text, FORMATO_NUM_1)
      
      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      ppAyuDet AYUDAT, 0
      ']
    End If
  End With
  Exit Sub
Err:
  gpErrores
  Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
  Dim dnContador As Integer
  Dim sSentencia As String
  
  'Llaves.
  txtLlave(0).Text = ""
  
  'Datos.
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  dtpDato.Value = Date
  pnCta_TpoTcb = TPOTCB_VTA
  For dnContador = 0 To 5
    txtDato(dnContador).Text = ""
  Next
  txtDato(3).Text = Format(0, FORMATO_NUM_2)
  txtDato(4).Text = Format(0, FORMATO_NUM_1)
  txtDato(5).Text = Format(0, FORMATO_NUM_1)
  
  txtDato(3).Tag = Format(0, FORMATO_NUM_2)
  txtDato(4).Tag = Format(0, FORMATO_NUM_1)
  txtDato(5).Tag = Format(0, FORMATO_NUM_1)

  'Ayudas.
  lblDatoDeta(0).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Byte
  
  'Datos.
  cboTpoMon.Enabled = tbHabilitar
  dtpDato.Enabled = tbHabilitar
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  cmdDatoAyud(0).Enabled = tbHabilitar
  lblDatoDeta(0).Enabled = tbHabilitar
End Sub

Private Sub cmdAuxiliar_Click()
  frmMAuxGrd.Show vbModal
  frmTConGrd.uorstTGAux.Requery
End Sub

Private Sub dtpDato_Validate(Cancel As Boolean)
      
  If Not (Month(dtpDato.Value) >= Val(gsMesAct) And Year(dtpDato.Value) >= Val(gsAnoAct)) Then
    MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operación", "The date does not correspond with operating period"), vbCritical
    dtpDato.SetFocus
    Cancel = True
    Exit Sub
  End If
  dtpDato.Tag = 0
  If (dtpDato.Tag <> dtpDato.Value) Then
    dtpDato.Tag = dtpDato.Value
    With frmTConGrd.uorstTGTCb
      If .RecordCount <> 0 Then
        .MoveFirst
        .Find "(FehTCb) = '" & Format(dtpDato.Value, "yyyy/mm/dd") & "'"
        ' [Adicional Agregado por Angel
        If .EOF Then
          MsgBox TEXT_9015, vbExclamation
          txtDato(3).Text = Format(0, FORMATO_NUM_2)
          txtDato(3).SetFocus
          Cancel = True
          Exit Sub
        Else
          txtDato(3).Text = Format(IIf(pnCta_TpoTcb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta), FORMATO_NUM_2)
        End If
        ']
      Else
         txtDato(3).Text = Format(0, FORMATO_NUM_2)
      End If
    End With
  End If

End Sub

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
