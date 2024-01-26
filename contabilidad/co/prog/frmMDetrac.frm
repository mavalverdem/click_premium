VERSION 5.00
Begin VB.Form frmMDetrac 
   Caption         =   "[Entidad]"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frarangos 
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame2 
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   1320
         Width           =   975
         Begin VB.CheckBox chkEstDetrac 
            Caption         =   "Activo"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   165
            Width           =   795
         End
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
         Top             =   240
         Width           =   915
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
         Left            =   1440
         TabIndex        =   2
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   660
         Width           =   4290
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
         Left            =   1440
         TabIndex        =   3
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1035
         Width           =   4290
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Detraccion:"
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
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   825
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
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   900
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
         Left            =   360
         TabIndex        =   13
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tasa : "
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
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1702
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1920
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
         Picture         =   "frmMDetrac.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmMDetrac.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmMDetrac.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmMDetrac.frx":049E
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
         Picture         =   "frmMDetrac.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmMDetrac.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMDetrac"
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
   
   With frmMDetracGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!coddetrac.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
      txtDato(gsIdioma - 1).MaxLength = .uorstMain!detdetrac.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain!detdetracx.DefinedSize
    ']
    'txtDato(2).MaxLength = .uorstMain!pctdetrac.DefinedSize
    txtDato(2).MaxLength = 6
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
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Detracción:", "Descripción:", "Traducción:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Detraction:", "Description:", "Translation:")
  Next nElemento
  chkEstDetrac.Caption = Choose(gsIdioma, "&Activo", "&Active")
  
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
  If Not (frmMDetracGrd.uorstMain.BOF And frmMDetracGrd.uorstMain.EOF) Then
   frmMDetracGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMDetracGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMDetracGrd.uorstMain, Me 'Cambiar Formulario de Grid.
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
   Dim dvFeCre, dvFeMdf As Variant
   On Error GoTo Err
   With frmMDetracGrd                     'Cambiar Formulario de Grid.
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
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain.Find "coddetrac='" & txtLlave(0).Text & "'"
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
  
   frmMDetracGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

'ini 2015-06-26 proceso detrac
'Private Sub cmdDatoAyud_Click(Index As Integer)
'   Select Case Index                   'Cambiar. Añadir índices.
'   Case 0
'      'txtDato(2).SetFocus
'   End Select
'   ppAyuBus Index
'End Sub
'fin 2015-06-26 proceso detrac

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

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
       KeyAscii = fValidInt(Index, KeyAscii, lblTexto(Index).Caption)
       Exit Sub
   End If

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
      With frmMDetracGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "coddetrac='" & txtLlave(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistro <> -1 Then .Bookmark = dvRegistro
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistro
         End If
      End With
      
'[REVISAR.
'ini 2015-06-26 proceso detrac
'   If Index = 0 Then
'      If Len(txtLlave(0).Text) = 1 Or Len(txtLlave(0).Text) = 3 Then
'         MsgBox Choose(gsIdioma, "El diario debe ser de 2 o 4 caracteres.", "The journal must be  2 or 4 characters."), vbExclamation
'         Cancel = True
'         Exit Sub
'      End If
'      If Len(Trim(txtLlave(0).Text)) = 4 Then
'         With frmMDetracGrd.uorstcodetrac
'            .Requery
'            .Find "coddetrac='" & Mid(txtLlave(0).Text, 1, 2) & "'"
'            If .EOF Then
'               MsgBox Choose(gsIdioma, "El diario ", "The journal ") & Mid(txtLlave(0).Text, 1, 2) & Choose(gsIdioma, " no existe.", " no exist."), vbCritical
'               Cancel = True
'               Exit Sub
'            End If
'         End With
'      End If
'   End If
'fin 2015-06-26 proceso detrac
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
   If KeyAscii <> 8 Then
      If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
         SendKeys "{TAB}"
      End If
   End If
    If Index = 2 Then
         KeyAscii = fValidDeci(Index, KeyAscii, lblTexto(Index + 1).Caption)
    End If
'        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
'            ' El 48 es 0 y el 57 es 9, 127 es SUPR y 8 es Backspace
'            Exit Sub
'        Else
'            KeyAscii = 0
'            MsgBox "Solo números para registrar " & lblTexto(Index + 1).Caption _
'            & " sin puntos, " & "ni comas, ni cualquier caracter especial!!"
'        End If
']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)

' Select Case Index    'Busca el dato en su tabla principal.
'   Case 2                           'Cambiar (añadir índices).
'      Cancel = ppAyuDet(0)
'      If Cancel Then Exit Sub
'   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                          'Cambiar (añadir índices).
'ini 2015-06-26 proceso detrac
'      modAyuBus.Lib_Cod "", txtDato(2).Text, 0, 0, Me.Top + frarangos.Top + txtDato(2).Top + txtDato(2).Height, Me.Left + frarangos.Left + txtDato(2).Left
'      txtDato(2).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(0).Caption = " " & frmOAyuBus.uvDato2
'fin 2015-06-26 proceso detrac
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
 Select Case tnIndex                 'Cambiar.
   Case 0
'ini 2015-06-26 proceso detrac
'       If txtDato(2).Text = "" Then
'         lblDatoDeta(0).Caption = ""
'         Exit Function
'      End If
'      With frmMDetracGrd.uorstCOLIB
'         .MoveFirst
'         .Find "tsadetrac='" & txtDato(2).Text & "'"
'         If .EOF Then
'            MsgBox TEXT_8006, vbExclamation
'            ppAyuDet = True
'         Else
'            lblDatoDeta(0).Caption = " " & IIf(IsNull(!DesLIB), "", !DesLIB)
'         End If
'      End With
'fin 2015-06-26 proceso detrac
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMDetracGrd.uorstMain
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !codemp = gsCodEmp
            '2015-06-26 proceso detrac .uorstMain!pdoano = gsAnoAct
            !coddetrac = txtLlave(0).Text
            '2015-06-26 proceso detrac .uorstMain!tsadetrac = txtDato(2).Text
         End If

        'Datos.
         !detdetrac = txtDato(gsIdioma - 1).Text
         !detdetracx = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
         !pctdetrac = IIf(Trim(txtDato(2).Text) = "", 0#, txtDato(2).Text)
         '!pctdetrac = txtDato(2).Text
         !estdetrac = IIf(chkEstDetrac.Value = vbChecked, ESTCCO_ACT, ESTCCO_INA)
     Else
        'Llaves.
         txtLlave(0).Text = !coddetrac
      
        'Datos.
         txtDato(gsIdioma - 1).Text = IIf(IsNull(!detdetrac), "", !detdetrac)
         txtDato(2 - gsIdioma).Text = IIf(IsNull(!detdetracx), "", !detdetracx)
         txtDato(2).Text = IIf(IsNull(!pctdetrac), "", !pctdetrac)
         chkEstDetrac.Value = IIf(!estdetrac = ESTCCO_ACT, vbChecked, vbUnchecked)
       
         '2015-06-26 proceso detrac ppAyuDet 0
         
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
   chkEstDetrac.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With

  'Ayudas.
   '2015-06-26 proceso detrac  lblDatoDeta(0).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer
   chkEstDetrac.Enabled = tbHabilitar

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

  'Ayudas.
  '2015-06-26 proceso detrac  cmdDatoAyud(0).Enabled = tbHabilitar
   '2015-06-26 proceso detrac lblDatoDeta(0).Enabled = tbHabilitar
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


