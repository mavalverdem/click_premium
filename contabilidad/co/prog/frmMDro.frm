VERSION 5.00
Begin VB.Form frmMDro 
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
      TabIndex        =   7
      Top             =   0
      Width           =   7335
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1035
         Width           =   4290
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   0
         Left            =   6960
         Picture         =   "frmMDro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1440
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
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diario:"
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
         TabIndex        =   17
         Top             =   240
         Width           =   450
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
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   900
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
         Left            =   360
         TabIndex        =   15
         Top             =   1095
         Width           =   855
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
         Left            =   2400
         TabIndex        =   14
         Top             =   1440
         Width           =   4515
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Libro : "
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
         TabIndex        =   13
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
      TabIndex        =   6
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
         Picture         =   "frmMDro.frx":01AA
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
         Picture         =   "frmMDro.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmMDro.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmMDro.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmMDro.frx":074A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmMDro.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMDro"
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
   
   With frmMDroGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!coddro.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
      txtDato(gsIdioma - 1).MaxLength = .uorstMain!DetDro.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain!DetDrox.DefinedSize
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
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diario:", "Descripci�n:", "Traducci�n:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journal:", "Description:", "Translation:")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']
   
End Sub

Private Sub Form_Activate()
 '[Busca detalle de c�digos            'Cambiar (habilitar/deshabilitar).
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
  If Not (frmMDroGrd.uorstMain.BOF And frmMDroGrd.uorstMain.EOF) Then
   frmMDroGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMDroGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMDroGrd.uorstMain, Me 'Cambiar Formulario de Grid.
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
   With frmMDroGrd                     'Cambiar Formulario de Grid.
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
       '[B�squeda de llave actual.     'Cambiar.
         .uorstMain.Find "CodDro='" & txtLlave(0).Text & "'"
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
  
   frmMDroGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. A�adir �ndices.
   Case 0
      'txtDato(2).SetFocus
   End Select
   ppAyuBus Index
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

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (a�adir �ndices).
'   Case 0
'      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
'         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
'      End If
'   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (a�adir �ndices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMDroGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodDro='" & txtLlave(0).Text & "'"
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
   If Index = 0 Then
      If Len(txtLlave(0).Text) = 1 Or Len(txtLlave(0).Text) = 3 Then
         MsgBox Choose(gsIdioma, "El diario debe ser de 2 o 4 caracteres.", "The journal must be  2 or 4 characters."), vbExclamation
         Cancel = True
         Exit Sub
      End If
      If Len(Trim(txtLlave(0).Text)) = 4 Then
         With frmMDroGrd.uorstCODro
            .Requery
            .Find "CodDro='" & Mid(txtLlave(0).Text, 1, 2) & "'"
            If .EOF Then
               MsgBox Choose(gsIdioma, "El diario ", "The journal ") & Mid(txtLlave(0).Text, 1, 2) & Choose(gsIdioma, " no existe.", " no exist."), vbCritical
               Cancel = True
               Exit Sub
            End If
         End With
      End If
   End If
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
'[ARREGLAR: Retrocede si Shift est� presionado.
   If KeyAscii <> 8 Then
      If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
         SendKeys "{TAB}"
      End If
   End If
']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
 Select Case Index    'Busca el dato en su tabla principal.
   Case 2                           'Cambiar (a�adir �ndices).
      Cancel = ppAyuDet(0)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                          'Cambiar (a�adir �ndices).
      modAyuBus.Lib_Cod "", txtDato(2).Text, 0, 0, Me.Top + frarangos.Top + txtDato(2).Top + txtDato(2).Height, Me.Left + frarangos.Left + txtDato(2).Left
      txtDato(2).Text = frmOAyuBus.uvDato1
      lblDatoDeta(0).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
 Select Case tnIndex                 'Cambiar.
   Case 0
      If txtDato(2).Text = "" Then
         lblDatoDeta(0).Caption = ""
         Exit Function
      End If
      With frmMDroGrd.uorstCOLIB
         .MoveFirst
         .Find "CodLIB='" & txtDato(2).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(0).Caption = " " & IIf(IsNull(!DesLIB), "", !DesLIB)
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMDroGrd
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain!codemp = gsCodEmp
            .uorstMain!pdoano = gsAnoAct
            .uorstMain!coddro = txtLlave(0).Text
            .uorstMain!codlib = txtDato(2).Text
         End If

        'Datos.
         .uorstMain!DetDro = txtDato(gsIdioma - 1).Text
         .uorstMain!DetDrox = IIf(txtDato(2 - gsIdioma).Text = "", Null, txtDato(2 - gsIdioma).Text)
         .uorstMain!codlib = txtDato(2).Text
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain!coddro
      
        'Datos.
         txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain!DetDro), "", .uorstMain!DetDro)
         txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain!DetDrox), "", .uorstMain!DetDrox)
         txtDato(2).Text = IIf(IsNull(.uorstMain!codlib), "", .uorstMain!codlib)
         
         ppAyuDet 0
         
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

  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
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


