VERSION 5.00
Begin VB.Form frmMPms 
   Appearance      =   0  'Flat
   Caption         =   "[Entidad]"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   7005
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
      Index           =   1
      Left            =   780
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   6720
      Picture         =   "frmMPms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   495
      Width           =   255
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   6300
      Picture         =   "frmMPms.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   855
      Width           =   255
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
      Index           =   2
      Left            =   780
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   6720
      Picture         =   "frmMPms.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   135
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1762
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
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
         Picture         =   "frmMPms.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmMPms.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmMPms.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMPms.frx":099C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmMPms.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmMPms.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   720
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
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
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
      TabIndex        =   18
      Top             =   540
      Width           =   675
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
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   17
      Top             =   480
      Width           =   5535
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
      Height          =   315
      Index           =   2
      Left            =   2580
      TabIndex        =   15
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Módulo:"
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
      TabIndex        =   13
      Top             =   900
      Width           =   555
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
      Height          =   315
      Index           =   0
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
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
      Top             =   180
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   6960
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frmMPms"
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
   
   With frmMPmsGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodUsr.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!CodEmp.DefinedSize
      txtLlave(2).MaxLength = .uorstMain!CodMdl.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
'      txtDato(0).MaxLength = .uorstMain!NomPms.DefinedSize
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
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
   If txtLlave(0).Text <> "" Then ppAyuDet AYULLA, 0
   If txtLlave(1).Text <> "" Then ppAyuDet AYULLA, 1
   If txtLlave(2).Text <> "" Then ppAyuDet AYULLA, 2
'   If txtDato(1).Text <> "" Then ppAyuDet AYUDAT, 1
 ']

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMPmsGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMPmsGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMPmsGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
'   cmdRetroceder.Enabled = False
'   cmdAvanzar.Enabled = False
'   cmdCorregir.Enabled = False
'   cmdGrabar.Enabled = True
'   cmdDeshacer.Enabled = True
'   upHabilitacion (True)
'
' '[Dato con el foco al corregir.       'Cambiar.
'   txtDato(0).SetFocus
' ']
End Sub

Public Sub cmdGrabar_Click()
'   On Error GoTo Err

 '[Verificación de datos.
'   If Len(Trim(txtDato(0).Text)) = 0 Or Len(Trim(txtDato(1).Text)) = 0 Then
'      If Len(Trim(txtDato(0).Text)) = 0 Then
'         MsgBox TEXT_6002, vbCritical
'         txtDato(0).SetFocus           'Descripción
'      Else
'         MsgBox TEXT_6002, vbCritical
'         txtDato(0).SetFocus           'Nombre
'      End If
'      Exit Sub
'   End If
 ']
   
   With frmMPmsGrd                     'Cambiar Formulario de Grid.
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
'            !FyHMdf = Now
         End If
         .Update
      End With
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
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
  
   frmMPmsGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2
      txtLlave(Index).SetFocus
'   Case 1
'      mskLlave(Index).SetFocus
   End Select
   ppAyuBus AYULLA, Index
End Sub

'Private Sub cmdDatoAyud_Click(Index As Integer)
'   Select Case Index                   'Cambiar. Añadir índices.
'   Case 0
'      txtDato(Index).SetFocus
'   Case 0, 1
'      mskDato(Index).SetFocus
'   End Select
'   ppAyuBus AYUDAT, Index
'End Sub

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus AYULLA, Index
'   End If
'End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYULLA, Index
   End If
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
'   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
   If pbValidada And cmdGrabar.Enabled Then cmdGrabar.SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 2
'      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
'         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
'      End If
'   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
   Case 0, 1, 2
      Cancel = ppAyuDet(AYULLA, Index)
      If Cancel Then Exit Sub
   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0 And Len(Trim(txtLlave(2).Text)) <> 0 Then
      With frmMPmsGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
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

'Private Sub txtDato_GotFocus(Index As Integer)
'   txtDato(Index).SelStart = 0
'   txtDato(Index).SelLength = txtDato(Index).MaxLength
'End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
'   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
'      SendKeys "{TAB}"
'   End If
']ARREGLAR.

 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

'Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus AYUDAT, Index
'   End If
'End Sub

'Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
'   On Error GoTo Err
'
'  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select
'
'  'Asigna 0 a campos numéricos si están vacíos.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select
'
'  'Busca el dato en su tabla principal.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
'
'   Exit Sub
'Err:
'   gpErrores
'End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Usr_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case 1                           'Cambiar (añadir índices).
         modAyuBus.Emp_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case 2                           'Cambiar (añadir índices).
         modAyuBus.Mdl_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
   Else
'      Select Case tnIndex
'      Case 0                           'Cambiar (añadir índices).
'         modAyuBus.Dro_Cod "Length(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
'         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'      End Select
   End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex                 'Cambiar.
      Case 0
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmMPmsGrd.uorstSGUsr
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodUsr='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !NomUsr
            End If
         End With
      Case 1
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmMPmsGrd.uorstTGEmp
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodEmp='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !RazEmp
            End If
         End With
      Case 2
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmMPmsGrd.uorstSGMdl
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodMdl='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !DetMdl
            End If
         End With
      End Select
   Else
'      Select Case tnIndex                 'Cambiar.
'      Case 0
'         If txtDato(tnIndex).Text = "" Then
'            lblDatoDeta(tnIndex).Caption = ""
'            Exit Function
'         End If
'         With frmTVtaGrd.uorstCODro
'            If .RecordCount > 0 Then .MoveFirst
'            .Find "CodDro='" & txtDato(tnIndex).Text & "'"
'            If .EOF Then
'               MsgBox TEXT_8006, vbExclamation
'               ppAyuDet = True
'            Else
'               lblDatoDeta(tnIndex).Caption = " " & !DetDro
'            End If
'         End With
'      End Select
   End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMPmsGrd.uorstMain           'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !CodUsr = txtLlave(0).Text
            !CodEmp = txtLlave(1).Text
            !CodMdl = txtLlave(2).Text
            !CodSis = gsCodSis
            !IndPms01 = 1
            !IndPms02 = 1
            !IndPms03 = 1
            !IndPms04 = 1
            !IndPms05 = 1
            !IndPms06 = 1
            !IndPms07 = 1
            !IndPms08 = 1
            !IndPms09 = 1
            !IndPms10 = 1
         End If

        'Datos.
'         !EstUsr = IIf(chkEstUsr.Value = vbChecked, ESTUSR_ACT, ESTUSR_INA)
'         !CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         !FehOpe = dtpFecha.Value
'         !CodMon = optMoneda(1).Value
'         !NomUsr = txtDato(0).Text
      Else
        'Llaves.
         txtLlave(0).Text = !CodUsr
         txtLlave(1).Text = !CodEmp
         txtLlave(2).Text = !CodMdl
      
        'Datos.
'         chkEstUsr.Value = IIf(!EstUsr = ESTUSR_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
'         dtpFecha.Value = !FehOpe
'         optMoneda(1).Value = !CodMon
'         txtDato(0).Text = IIf(IsNull(!NomUsr), "", !NomUsr)
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
   txtLlave(1).Text = ""
   txtLlave(2).Text = ""

  'Datos.
'   chkEstUsr.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
'   With txtDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Text = ""
'      Next
'   End With

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
   lblLlaveDeta(1).Caption = ""
   lblLlaveDeta(2).Caption = ""
'   lblDatoDeta(2).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
'   With txtDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With

  'Ayudas.
'   cmdDatoAyud(2).Enabled = tbHabilitar
'   lblDatoDeta(2).Enabled = tbHabilitar
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


