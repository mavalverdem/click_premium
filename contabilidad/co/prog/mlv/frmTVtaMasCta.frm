VERSION 5.00
Begin VB.Form frmTVtaMasCta 
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
      Index           =   5
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
      Index           =   4
      Left            =   7290
      Picture         =   "frmTVtaMasCta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
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
      Index           =   4
      Left            =   960
      MaxLength       =   11
      TabIndex        =   4
      Top             =   480
      Width           =   1275
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
      TabIndex        =   7
      Top             =   840
      Width           =   6525
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2130
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   22
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
         Picture         =   "frmTVtaMasCta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "frmTVtaMasCta.frx":02F4
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmTVtaMasCta.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmTVtaMasCta.frx":04F8
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmTVtaMasCta.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmTVtaMasCta.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   15
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
      Index           =   2
      Left            =   3600
      TabIndex        =   14
      Top             =   2670
      Width           =   1815
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
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   2670
      Width           =   1815
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7290
      Picture         =   "frmTVtaMasCta.frx":0996
      Style           =   1  'Graphical
      TabIndex        =   21
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
      Index           =   4
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
      Caption         =   "M.E."
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
      Index           =   6
      Left            =   3240
      TabIndex        =   13
      Top             =   2730
      Width           =   300
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "M.N."
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
      Index           =   5
      Left            =   840
      TabIndex        =   11
      Top             =   2730
      Width           =   315
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
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
      Width           =   570
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
Attribute VB_Name = "frmTVtaMasCta"
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
   
   With frmTVtaGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtDato(0).MaxLength = .uorstCOVtaDocCta!CodCta.DefinedSize
    ']
    '[Datos                            'Cambiar.
      txtDato(4).MaxLength = 11
'ini 2016-04-26 corrige tamaño
'      txtDato(3).MaxLength = 250
'      txtDato(5).MaxLength = 250
      txtDato(3).MaxLength = 71
      txtDato(5).MaxLength = 71
'ini 2016-04-26 corrige tamaño
      txtDato(1).MaxLength = 14
      txtDato(2).MaxLength = 14
   End With
'//Raul 12/12/2003
'// ¿?
'   cmdRetroceder.Enabled = (Not pbNuevo)
'   cmdCorregir.Enabled = (Not pbNuevo)
'   cmdAvanzar.Enabled = (Not pbNuevo)
'//
'/// Angel 12/12/2003
'// Habilitar los botones cuando es Nuevo
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
'   cmdSalir.Enabled = pbNuevo
   cmdCorregir.Enabled = (Not pbNuevo)
   cmdRetroceder.Enabled = (Not pbNuevo)
   cmdAvanzar.Enabled = (Not pbNuevo)
'//   cmdGrabar.Enabled = False
'//   cmdDeshacer.Enabled = False
'///
   upHabilitacion pbNuevo
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(7, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuenta :", "Auxiliar :", "Glosa :", "Traducción :", "Importe :", "MN", "ME")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Account :", "Auxiliary :", "Gloss  :", "Translation :", "Amount :", "NC", "FC")
  Next nElemento
  cmdGrabar.Caption = Choose(gsIdioma, "&Aceptar", "&Accept")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, False, True, True, aLabel
  ']
  
   'frmTVta.pglodet_len = Len(Trim(txtDato(3).Text)) '2016-02-23 control tamaño de glosa en formato
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
'///Angel 12/12/2003
'/// Se agrego habilitacion del boton salir
   cmdSalir.Enabled = True
'///
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(4).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmTVtaGrd                     'Cambiar Formulario de Grid.
'      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstCOVtaDocCta.AddNew
      End If
      upDatosDesconectados 0
      With .uorstCOVtaDocCta
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
'      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstCOVtaDocCta.Requery
         .upDatosGrid
''       '[Búsqueda de llave actual.     'Cambiar.
''         .uorstCOVtaDocCta.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
''       ']
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
  
'   frmTVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
   Case 4
      txtDato(Index).SetFocus
      
'   Case 3
'      mskDato(Index).SetFocus
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
 
 '[Convierte a mayúsculas.
'   If Index = 0 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
 
''ini 2016-02-23 control tamaño de glosa en formato
' If Index = 3 Then
'    If KeyAscii = 8 Then 'si es tecla retroceso
'        frmTVta.pglodet_len = frmTVta.pglodet_len - 1
'    Else
'        frmTVta.pglodet_len = frmTVta.pglodet_len + 1
'    End If
'    If frmTVta.pglodet_len > frmTVta.pglodet_len_max Then
'       frmTVta.pglodet_len = frmTVta.pglodet_len - 1
'       MsgBox ("No puede ingresar mas de " & Str(frmTVta.pglodet_len_max) & " caracteres")
'       KeyAscii = 0
'    End If
' End If
''fin 2016-02-23 control tamaño de glosa en formato

End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
   If Index = 0 Then
   
   ElseIf Index > 0 Then
      With frmTVta
         If .chkMonedaActiva.Value = vbChecked Then
            If Index = 1 Then
               If .cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                  txtDato(2).Text = Format(gfRedond(CDec(txtDato(1).Text) / CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               ElseIf CDec(txtDato(2).Text) = 0 Then
                  txtDato(2).Text = Format(gfRedond(CDec(txtDato(1).Text) / CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               End If
            End If
            If Index = 2 Then
               If .cboTpoMon.ListIndex = TPOMON_EXT_IND Then
                  txtDato(1).Text = Format(gfRedond(CDec(txtDato(2).Text) * CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               ElseIf CDec(txtDato(1).Text) = 0 Then
                  txtDato(1).Text = Format(gfRedond(CDec(txtDato(2).Text) * CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               End If
            End If
         End If
      End With
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistroActual As Variant
  
  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 1, 21 To 28                    'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

  'Asigna 0 a campos numéricos si están vacíos.
   Select Case Index
   Case 1, 2                           'Cambiar (añadir índices).
      If txtDato(Index).Text = "" Then
         txtDato(Index).Text = 0
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
      '[
      With frmTVtaGrd.uorstCOVtaDocCta
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
'///Angel 12/12/2003
'/// Solo se agrego el la propiedad Text en TxtDato(0) al final de la instruccion de busqueda
             .Find "cLlave2='" & frmTVta.txtLlave(0).Text & frmTVta.txtLlave(1).Text & frmTVta.txtLlave(2).Text & frmTVtaMasGrd.unIndice & ps_OrdCuenta & txtDato(0).Text & "'"
'///
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
     'cambio teo
   Case 4
       If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      '[
      With frmTVtaGrd.uorstCOVtaDocCta
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
'///Angel 12/12/2003
'/// Solo se agrego el la propiedad Text en TxtDato(0) al final de la instruccion de busqueda
             .Find "cLlave2='" & frmTVta.txtLlave(0).Text & frmTVta.txtLlave(1).Text & frmTVta.txtLlave(2).Text & frmTVtaMasGrd.unIndice & ps_OrdCuenta & txtDato(4).Text & "'"
'///
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
   Case 4                           'Cambiar (añadir índices).
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
      With frmTVtaGrd.uorstCoCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
   Case 4
    If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTVtaGrd.uorstTGAux
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

   With frmTVtaGrd.uorstCOVtaDocCta    'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !codtdc = frmTVta.txtLlave(0).Text
            !serdoc = frmTVta.txtLlave(1).Text
            !nrodoc = frmTVta.txtLlave(2).Text
            !tpocnc = frmTVtaMasGrd.unIndice
            !orden = ps_OrdCuenta
         End If

        'Datos.
         !CodCta = IIf(txtDato(0).Text = "", Null, txtDato(0))
         !codruc = IIf(txtDato(4).Text = "", Null, txtDato(4))
         !glodet0 = Left(IIf(txtDato(Choose(gsIdioma, 3, 5)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 5)).Text), 250)
         !glodet1 = Mid(IIf(txtDato(Choose(gsIdioma, 3, 5)).Text = "" Or Len(txtDato(Choose(gsIdioma, 3, 5)).Text) <= 250, Null, txtDato(Choose(gsIdioma, 3, 5)).Text), 251)
         !glodet0x = Left(IIf(txtDato(Choose(gsIdioma, 5, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 5, 3)).Text), 250)
         !glodet1x = Mid(IIf(txtDato(Choose(gsIdioma, 5, 3)).Text = "" Or Len(txtDato(Choose(gsIdioma, 5, 3)).Text) <= 250, Null, txtDato(Choose(gsIdioma, 5, 3)).Text), 251)
         !impcta_mn = CDec(txtDato(1).Text)
         !impcta_me = CDec(txtDato(2).Text)
      Else
        'Datos.
         txtDato(0).Text = IIf(IsNull(!CodCta), "", !CodCta)
         txtDato(4).Text = IIf(IsNull(!codruc), "", !codruc)
         txtDato(Choose(gsIdioma, 3, 5)).Text = IIf(IsNull(!glodet), "", !glodet)
         txtDato(Choose(gsIdioma, 5, 3)).Text = IIf(IsNull(!glodetx), "", !glodetx)
         txtDato(1).Text = Format(!impcta_mn, FORMATO_NUM_1)
         txtDato(2).Text = Format(!impcta_me, FORMATO_NUM_1)
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
'   cboTpoMon.ListIndex = TPOMON_NAC_IND
'   chkEstado.Value = vbChecked
'   dtpDato(3).Value = Date
'   optTpoMon(1).Value = True
   txtDato(0).Text = ""
   txtDato(4).Text = ""
   txtDato(3).Text = ""
   txtDato(5).Text = ""
   txtDato(1).Text = Format(0, FORMATO_NUM_1)
   txtDato(2).Text = Format(0, FORMATO_NUM_1)
   ps_OrdCuenta = "00"
'///Angel 22/12/2003
'///Envio de valor desde la pantalla de inicio de registro
   If pbNuevo Then
'ini 2016-04-26 corrige tamaño
'      txtDato(3).Text = frmTVta.txtDato(3).Text
      txtDato(3).Text = Left(frmTVta.txtDato(3).Text, 75)
'ini 2016-04-26 corrige tamaño
      txtDato(5).Text = Left(frmTVta.txtDato(37).Text, 75) '12-09-2016 corrige glosa vta
      If frmTVta.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
         txtDato(1).Text = Format(frmTVta.txtDato(frmTVtaMasGrd.unIndice + 4).Text, FORMATO_NUM_1)
      Else
         txtDato(2).Text = Format(frmTVta.txtDato(frmTVtaMasGrd.unIndice + 11).Text, FORMATO_NUM_1)
      End If
      With frmTVtaGrd.uorstCOVtaDocCta
         If Not .EOF And .RecordCount > 0 Then
            .MoveFirst
            Do
               If frmTVta.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                  txtDato(1).Text = Format(CDec(txtDato(1).Text) - !impcta_mn, FORMATO_NUM_1)
               Else
                  txtDato(2).Text = Format(CDec(txtDato(2).Text) - !impcta_me, FORMATO_NUM_1)
               End If
               ps_OrdCuenta = IIf(!orden > ps_OrdCuenta, !orden, ps_OrdCuenta)
               .MoveNext
            Loop Until .EOF
            .MoveFirst
         End If
      End With
   End If
'///
ps_OrdCuenta = gfCeros(ps_OrdCuenta, 2, 1, "0")
   
  'Ayudas.
   For dnContador = 0 To 0
      lblDatoDeta(dnContador).Caption = ""
   Next
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
'/// Angel 12/12/2003
'/// Se agrego condicion para no permitir la correccion del dato Cuenta, solo para los importes
      For dnContador = 0 To .Count - 1
         If dnContador = 0 Then
            .Item(dnContador).Enabled = pbNuevo
         Else
            .Item(dnContador).Enabled = tbHabilitar
         End If
'         .Item(dncontador).Enabled = tbHabilitar
      Next
'///
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

End Property
