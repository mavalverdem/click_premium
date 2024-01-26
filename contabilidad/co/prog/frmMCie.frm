VERSION 5.00
Begin VB.Form frmMCie 
   Caption         =   "[Entidad]"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   0
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
      TabIndex        =   1
      Top             =   780
      Width           =   5475
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1545
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1620
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
         Picture         =   "frmMCie.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmMCie.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmMCie.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMCie.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmMCie.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmMCie.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro. de Asiento:"
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
      Left            =   60
      TabIndex        =   8
      Top             =   180
      Width           =   1170
   End
   Begin VB.Label Label2 
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
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   6520
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMCie"
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
   
   With frmMCieGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain_0!NroCie.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
      txtDato(0).MaxLength = .uorstMain_0!DetCie.DefinedSize
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
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
   If Not frmMCieGrd.uorstMain_0.EOF Then
      If frmMCieGrd.uorstMain_0.EditMode <> adEditNone Then frmMCieGrd.uorstMain_0.CancelUpdate 'Cambiar Formulario de Grid.
   End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMCieGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMCieGrd.uorstMain_0, Me 'Cambiar Formulario de Grid.
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

   With frmMCieGrd                     'Cambiar Formulario de Grid.
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
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain_0.Find "NroCie='" & txtLlave(0).Text & "'"
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
  
   frmMCieGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
   Case 0
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
      End If
   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMCieGrd.uorstMain_0        'Cambiar Formulario de Grid.
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "NroCie='" & txtLlave(0).Text & "'"
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
']ARREGLAR.

 '[Convierte a mayúsculas.
   If Index = 1 Then                   'Cambiar (añadir índices).
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
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

   With frmMCieGrd                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain_0!NroCie = txtLlave(0).Text
         End If

        'Datos.
'         uorstMain!EstCCo = IIf(chkEstado.Value = vbChecked, ESTCCO_ACT, ESTCCO_INA)
'         uorstMain!CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         uorstMain!FehOpe = dtpFecha.Value
'         uorstMain!CodMon = optMoneda(1).Value
         .uorstMain_0!DetCie = txtDato(0).Text
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain_0!NroCie
      
        'Datos.
'         chkEstado.Value = IIf(uorstMain!EstCCo = ESTCCO_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(uorstMain!CodSoc), "", uorstMain!CodSoc)
'         dtpFecha.Value = uorstMain!FehOpe
'         optMoneda(1).Value = uorstMain!CodMon
         txtDato(0).Text = IIf(IsNull(.uorstMain_0!DetCie), "", .uorstMain_0!DetCie)
      
        'Busca detalle de códigos.
'         ppAyuDet 0
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

