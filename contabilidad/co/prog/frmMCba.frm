VERSION 5.00
Begin VB.Form frmMCba 
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1920
      TabIndex        =   18
      Top             =   1680
      Width           =   4215
   End
   Begin VB.ComboBox cbotpocta 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2040
      Width           =   1980
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmMCba.frx":0000
      Left            =   1920
      List            =   "frmMCba.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1980
   End
   Begin VB.Frame fraRangos 
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   ".."
         Height          =   280
         Index           =   0
         Left            =   6600
         Picture         =   "frmMCba.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtllave 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Bancos"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   160
         Width           =   540
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
         Left            =   750
         TabIndex        =   14
         Top             =   360
         Width           =   5820
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1800
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
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
         Picture         =   "frmMCba.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmMCba.frx":02F8
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1200
         Picture         =   "frmMCba.frx":03FA
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmMCba.frx":04FC
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmMCba.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMCba.frx":07F0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   360
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Interbancaria :"
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
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   0
      X2              =   6780
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cuenta :"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1170
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Numero Cuenta :"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1200
   End
End
Attribute VB_Name = "frmMCba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
 
Private Const BM_SETSTATE = &HF3
Private Const WM_LBUTTONDOWN = &H201 ' botón izquierdo abajo
Private Const WM_LBUTTONUP = &H202 ' izquierdo arriba
Private Const WM_LBUTTONDBLCLK As Long = &H203 ' izquierdo doble click

' enviar pulsación de mouse al Hwnd indicado
Sub Enviar_Pulsacion(Handle As Long)
    Call SendMessage(Handle, BM_SETSTATE, 0, ByVal 0&)
    Call SendMessage(Handle, WM_LBUTTONDOWN, 0, ByVal 0&)
    Call SendMessage(Handle, WM_LBUTTONUP, 0, ByVal 0&)
    Call SendMessage(Handle, BM_SETSTATE, 1, ByVal 0&)
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
 ppAyuBus Index
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0                           'Cambiar (añadir índices).
    modAyuBus.Bco_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + fraRangos.Left + txtLlave(tnIndex).Left
    txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Sub Form_Load()
   
   txtDato(0).MaxLength = 20
   txtDato(1).MaxLength = 20
   
   pbValidada = False

   Me.KeyPreview = True
   
   With cboTpoMon
      .AddItem TPOMON_NAC_TXT_2, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_2, TPOMON_EXT_IND
   End With
   
   With cboTpoCta
      .AddItem TPOCTA_COR_TXT_2, TPOCTA_COR_IND
      .AddItem TPOCTA_AHO_TXT_2, TPOCTA_AHO_IND
      .AddItem TPOCTA_MAE_TXT_2, TPOCTA_MAE_IND
    End With
   
   With frmMCbaGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
        txtLlave(0).MaxLength = .uorstMain!codbco.DefinedSize
    ']
    '[Datos                            'Cambiar.
        txtDato(0).MaxLength = .uorstMain!nroctacte.DefinedSize
    ']
     If pbNuevo Then
     Else
        If .uorstMain!tpomon = TPOMON_NAC_IND Then
            cboTpoMon.Text = TPOMON_NAC_TXT_2
        Else
            cboTpoMon.Text = TPOMON_EXT_TXT_2
        End If
        If .uorstMain!TpoCTA = TPOCTA_AHO_IND Then
            cboTpoCta.Text = TPOCTA_AHO_TXT_2
        ElseIf .uorstMain!TpoCTA = TPOCTA_COR_IND Then
            cboTpoCta.Text = TPOCTA_COR_TXT_2
        Else
            cboTpoCta.Text = TPOCTA_MAE_TXT_2
        End If
     End If
   End With
   
   If pbNuevo Then
      
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
      
      cmdGrabar.Enabled = True
      cmdDeshacer.Enabled = False
      upHabilitacion True
      pbValidada = True
      
   Else
      
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      upHabilitacion False
   
   End If
   
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Banco :", "Moneda :", "N-Cta.Cte :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Bank :", "Money :", "N-Cta.Cte :")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']
   
  
End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
    If txtLlave(0).Text <> "" Then ppAyuDet 0
 ']

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err
 If Not (frmMCbaGrd.uorstMain.BOF And frmMCbaGrd.uorstMain.EOF) Then
     frmMCbaGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
Err:
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMCbaGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMCbaGrd.uorstMain, Me 'Cambiar Formulario de Grid.
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
  
  If Trim(txtLlave(0).Text) = "" Then MsgBox TEXT_8005, vbExclamation: txtLlave(0).SetFocus: Exit Sub
  If Trim(cboTpoMon.Text) = "" Then MsgBox TEXT_8005, vbExclamation: cboTpoMon.SetFocus: Exit Sub
  If Trim(txtDato(0).Text) = "" Then MsgBox TEXT_8005, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If Trim(cboTpoCta.Text) = "" Then MsgBox TEXT_8005, vbExclamation: cboTpoCta.SetFocus: Exit Sub
 
   On Error GoTo Err

   With frmMCbaGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain
         If pbNuevo Then
            !usrcre = gsAbvUsr
            !fyhcre = Now
            .Update
         Else
            '!UsrMdf = gsAbvUsr
            '!FyHMdf = Now
            
         Dim Rstupdate As ADODB.Recordset
         Dim sql As String
        
         Set Rstupdate = New ADODB.Recordset
      
         sql = "update coctaban set nroctacte='" & txtDato(0).Text & "', tpocta='" & Choose(cboTpoCta.ListIndex + 1, TPOCTA_COR, TPOCTA_AHO, TPOCTA_MAE) & "',nrocci='" & txtDato(1).Text & "'"
         sql = sql & " where codemp='" & gsCodEmp & "' "
         sql = sql & " and codaux='" & CtaAuxiliar & "' "
         sql = sql & " and codbco='" & txtLlave(0).Text & "'"
         sql = sql & " and tpomon='" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT) & "'"
     
         Rstupdate.Open sql, frmMCbaGrd.uocnnMain, adOpenStatic, adLockOptimistic
                       
        End If
         
      End With
       
      '.uorstCCCfg.Update
      
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         
         .uorstMain.Requery
         .ppDatosGrid
       
       '[Búsqueda de llave actual.     'Cambiar.
        .uorstMain.Find "codbco='" & txtLlave(0).Text & "'"
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

   'Call Enviar_Pulsacion(frmMCbaGrd.cmdRefrescar.hwnd)
    
   Unload Me
   Exit Sub
    
Err:
   'gpErrores
   frmMCbaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)

'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(txtLlave(Index))) + 1 = txtLlave(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.

 '[Convierte a mayúsculas.
   If Index = 1 Then                   'Cambiar (añadir índices).
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
 ']
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
      ppAyuBus Index
    End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1           'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMCbaGrd
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
          .uorstMain!codemp = gsCodEmp
          .uorstMain!codaux = CtaAuxiliar
          .uorstMain!codbco = txtLlave(0).Text
          .uorstMain!tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
          .uorstMain!TpoCTA = Choose(cboTpoCta.ListIndex + 1, TPOCTA_AHO, TPOCTA_COR, TPOCTA_MAE)
         End If
        'Datos.
         .uorstMain!nroctacte = txtDato(0).Text
         .uorstMain!nrocci = txtDato(1).Text
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain!codbco
         
         cboTpoMon.ListIndex = IIf(.uorstMain!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         cboTpoCta.ListIndex = IIf(.uorstMain!TpoCTA = TPOCTA_AHO, TPOCTA_AHO_IND, IIf(.uorstMain!TpoCTA = TPOCTA_COR, TPOCTA_COR_IND, TPOCTA_MAE_IND))
         
        'Datos.
         txtDato(0).Text = IIf(IsNull(.uorstMain!nroctacte), "", .uorstMain!nroctacte)
         txtDato(1).Text = IIf(IsNull(.uorstMain!nrocci), "", .uorstMain!nrocci)
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
  
   With txtDato
      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Text = ""
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
      '   .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

  'Ayudas.
  'cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
    
      If txtLlave(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCbaGrd.uorstTGBCO
         .MoveFirst
         .Find "Codbco='" & txtLlave(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !detbco
         End If
      End With
   End Select
End Function

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





