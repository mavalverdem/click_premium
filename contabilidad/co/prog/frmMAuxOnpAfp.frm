VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMAuxOnpAfp 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1800
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   11
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
         Picture         =   "frmMAuxOnpAfp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   60
         Visible         =   0   'False
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
         Picture         =   "frmMAuxOnpAfp.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   345
         Visible         =   0   'False
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
         Picture         =   "frmMAuxOnpAfp.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
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
         Picture         =   "frmMAuxOnpAfp.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
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
         Picture         =   "frmMAuxOnpAfp.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmMAuxOnpAfp.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   12
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
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   4215
   End
   Begin VB.Frame fraRangos 
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtllave 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   ".."
         Height          =   280
         Index           =   0
         Left            =   6600
         Picture         =   "frmMAuxOnpAfp.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
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
         Height          =   285
         Index           =   0
         Left            =   750
         TabIndex        =   5
         Top             =   360
         Width           =   5820
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Entidad de Pensión"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   165
         Width           =   1380
      End
   End
   Begin VB.ComboBox cboFlagComision 
      Height          =   315
      ItemData        =   "frmMAuxOnpAfp.frx":0996
      Left            =   1920
      List            =   "frmMAuxOnpAfp.frx":0998
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   1980
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   66977793
      CurrentDate     =   37102
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
      Caption         =   "Fecha Nacimiento:"
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
      TabIndex        =   10
      Top             =   1680
      Width           =   1320
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Código Cussp :"
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
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Comisiòn:"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1035
   End
End
Attribute VB_Name = "frmMAuxOnpAfp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub puorstOnp_refresh()
    With frmMAux.uorstOnp
    If Not .EOF Then
        '.MoveFirst
        txtLlave(0).Text = !CodAfp
        txtDato(0).Text = !numeroafp
        cboFlagComision.ListIndex = IIf(IsNull(!Flagcomision), 0, !Flagcomision)
        dtpDato(0).Value = !Fecnacimiento
    End If
    End With
End Sub
Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
    If tsTipo = AYULLA Then
        Select Case tnIndex                 'Cambiar.
        Case 0
            If txtLlave(tnIndex).Text = "" Then
              lblLlaveDeta(tnIndex).Caption = ""
              Exit Function
            End If
            With frmMAuxGrd.uorstCoEntidadPen
              If .RecordCount > 0 Then .MoveFirst
                .Find "Codafp='" & txtLlave(tnIndex).Text & "'"
              If .EOF Then
                MsgBox TEXT_8006, vbExclamation
                ppAyuDet = True
              Else
                lblLlaveDeta(tnIndex).Caption = " " & IIf(IsNull(!Desafp), "", !Desafp)
              End If
            End With
        End Select
    Else
    End If
End Function

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
    If tsTipo = AYULLA Then
      Select Case tnIndex
       Case 0                           'Cambiar (añadir índices).
        modAyuBus.OnoAfp_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
        txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
        lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
    Else
    End If
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtLlave(Index).SetFocus
'   Case 1
'      mskLlave(Index).SetFocus
   End Select
   ppAyuBus AYULLA, Index

End Sub

Private Sub cmdDeshacer_Click()
    frmMAux.puorstOnp_Insert
    puorstOnp_refresh
End Sub

Private Sub cmdSalir_Click()
    With frmMAux.uorstOnp
        If .EOF Then
            .AddNew
        End If
        .Fields("CodAfp") = txtLlave(0).Text
        .Fields("NumeroAfp") = txtDato(0).Text
        '2014-08-25 error flag.Fields("FlagComision") = Choose(cboFlagComision.ListIndex + 1, INDCOMI_FLU, INDCOMI_MIX)
        '2014-08-27 error -1 con sql
        .Fields("FlagComision") = IIf(cboFlagComision.ListIndex = -1, 0, cboFlagComision.ListIndex)
        '.Fields("FlagComision") = cboFlagComision
        .Fields("FecNacimiento") = dtpDato(0).Value
        .Update
        If Len(Trim(.Fields("CodAfp"))) = 0 Then
            MsgBox Choose(gsIdioma, "Ponga el codigo de AFP, si no perdera todos los datos", "Put the code AFP, if not lose all data"), vbExclamation
        End If
    End With
   Unload Me
End Sub

Private Sub Form_Load()

    With cboFlagComision
      .AddItem TPOCOMI_MIXTA_TXT, INDCOMI_MIX
      .AddItem TPOCOMI_FLUJO_TXT, INDCOMI_FLU
    End With
    
    puorstOnp_refresh
    
    txtLlave(0).MaxLength = frmMAux.uorstOnp!CodAfp.DefinedSize
    txtDato(0).MaxLength = frmMAux.uorstOnp!numeroafp.DefinedSize
    
'    With frmMAux.uorstOnp
'    If Not .EOF Then
'        '.MoveFirst
'        txtllave(0).Text = !CodAfp
'        txtDato(0).Text = !numeroafp
'        cboFlagComision.ListIndex = !Flagcomision
'        dtpDato(0).Value = !Fecnacimiento
'    End If
'    End With
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

Private Sub txtllave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYULLA, Index
   End If
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
 '[Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 0, 1, 2                        'Cambiar (añadir índices).
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
      End If
   End Select
 ']
 '[Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYULLA, Index)
      If Cancel Then Exit Sub
   End Select
 ']
   Exit Sub
Err:
   gpErrores

End Sub
