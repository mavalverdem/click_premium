VERSION 5.00
Begin VB.Form frmMEFi 
   Caption         =   "[Entidad]"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   2
      Left            =   6090
      Picture         =   "frmMEFi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1500
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   2
      Left            =   1020
      TabIndex        =   7
      Top             =   1500
      Width           =   520
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
      Top             =   1125
      Width           =   5355
   End
   Begin VB.CheckBox chkFmtCnv 
      Alignment       =   1  'Right Justify
      Caption         =   "Formato CONASEV"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   50
      TabIndex        =   9
      Top             =   1890
      Width           =   1875
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
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   315
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
      Top             =   735
      Width           =   5355
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1470
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2190
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
         Picture         =   "frmMEFi.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmMEFi.frx":02F4
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmMEFi.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmMEFi.frx":04F8
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmMEFi.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmMEFi.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   10
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
      Height          =   280
      Index           =   2
      Left            =   1545
      TabIndex        =   8
      Top             =   1500
      Width           =   4545
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Proyecto :"
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
      TabIndex        =   6
      Top             =   1545
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducci�n:"
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
      Top             =   1185
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Estado Financiero:"
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
      Width           =   1335
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Descripci�n:"
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
      Top             =   795
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   6360
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMEFi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
'Private porstTGTPv As ADODB.Recordset
']

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmMEFiGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain_0!CodEfi.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
      txtDato(gsIdioma - 1).MaxLength = .uorstMain_0!DetEFi.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain_0!DetEFix.DefinedSize
      txtDato(2).MaxLength = .uorstMain_0!coddpe.DefinedSize
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
  ReDim aLabel(4, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Estado Financiero:", "Descripci�n:", "Traducci�n:", "Proyecto :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Financial Statement:", "Description:", "Translation:", "Project :")
  Next nElemento
  chkFmtCnv.Caption = Choose(gsIdioma, "Formato CONASEV", "Format CONASEV")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
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

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMEFiGrd.uorstMain_0.BOF And frmMEFiGrd.uorstMain_0.EOF) Then
   frmMEFiGrd.uorstMain_0.CancelUpdate 'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMEFiGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMEFiGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
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
  
  If txtDato(2).Text = "" Then MsgBox TEXT_8005, vbInformation: txtDato(2).SetFocus: Exit Sub
   
   With frmMEFiGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain_0.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain_0
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
            !FyHMdf = Now
         End If
         .Update
      End With
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain_0.Requery
         .upDatosGrid 0
       '[B�squeda de llave actual.     'Cambiar.
         .uorstMain_0.Find "CodEFi='" & txtLlave(0).Text & "'"
       ']
         cmdGrabar.Enabled = False
         upHabilitacion False
   
         upDatosPredeterminados
       '[Llave con el foco al a�adir.  'Cambiar.
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
  
   frmMEFiGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. A�adir �ndices.
   Case 2
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
 '[Convierte a may�sculas.
'   If Index = 1 Then                   'Cambiar (a�adir �ndices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (a�adir �ndices).
   Case 0
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
      End If
   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (a�adir �ndices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMEFiGrd.uorstMain_0        'Cambiar Formulario de Grid.
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodEFi='" & txtLlave(0).Text & "'"
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
'[ARREGLAR: Retrocede si Shift est� presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.

 '[Convierte a may�sculas.
   If Index = 1 Then                   'Cambiar (a�adir �ndices).
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
 ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 12                             'Cambiar (a�adir �ndices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

  'Asigna 0 a campos num�ricos si est�n vac�os.
'   Select Case Index
'   Case 2                              'Cambiar (a�adir �ndices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select

  'Busca el dato en su tabla principal.
  Select Case Index
   Case 2                             'Cambiar (a�adir �ndices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
  Exit Sub

Err:
  gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 2                              'Cambiar (a�adir �ndices).
    modAyuBus.DPe_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(coddpe)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 2
    If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
    With frmMEFiGrd.uorstCoDPe
      .MoveFirst
      .Find "coddpe='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detdpe), "", !detdpe)
      End If
    End With
  End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMEFiGrd                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain_0!codemp = gsCodEmp
            .uorstMain_0!pdoano = gsAnoAct
            .uorstMain_0!CodEfi = txtLlave(0).Text
         End If

        'Datos.
         .uorstMain_0!IndCnv = IIf(chkFmtCnv.Value = vbChecked, INDCNV_VER, INDCNV_FAL)
         .uorstMain_0!DetEFi = txtDato(gsIdioma - 1).Text
         .uorstMain_0!DetEFix = txtDato(2 - gsIdioma).Text
         .uorstMain_0!coddpe = txtDato(2).Text
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain_0!CodEfi
      
        'Datos.
         chkFmtCnv.Value = IIf(.uorstMain_0!IndCnv = INDCNV_VER, vbChecked, vbUnchecked)
         txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain_0!DetEFi), "", .uorstMain_0!DetEFi)
         txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain_0!DetEFix), "", .uorstMain_0!DetEFix)
         txtDato(2).Text = IIf(IsNull(.uorstMain_0!coddpe), "", .uorstMain_0!coddpe)
        'Busca detalle de c�digos.
         ppAyuDet 2
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
   chkFmtCnv.Value = vbUnchecked
'   chkEstado.Value = vbChecked
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
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   chkFmtCnv.Enabled = tbHabilitar

  'Ayudas.
   cmdDatoAyud(2).Enabled = tbHabilitar
End Sub

'[C�digo propio del formulario.

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

