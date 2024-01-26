VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTCprMasGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   4875
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox TxtDetaGrd 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   7095
   End
   Begin VB.TextBox TxtDetaGrd 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7095
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   1635
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2884
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cuenta"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7095
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Re&frescar"
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
         Left            =   2160
         Picture         =   "frmTCprMasGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
      Begin VB.Frame fraBuscar 
         Caption         =   "&Buscar por [Columna]"
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
         Left            =   3650
         TabIndex        =   7
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   200
            Width           =   2415
         End
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
         Left            =   6375
         Picture         =   "frmTCprMasGrd.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   0
         Picture         =   "frmTCprMasGrd.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
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
         Left            =   1440
         Picture         =   "frmTCprMasGrd.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   2880
         Picture         =   "frmTCprMasGrd.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdRevisar 
         Caption         =   "&Revisar"
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
         Left            =   720
         Picture         =   "frmTCprMasGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrCCosto 
      Height          =   1635
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2884
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Centro de Costo"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTCprMasGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private paOpciones As Variant
Private pnColumnaOrd As Integer

'[Propio del formulario.
Public unIndice As Byte
Private psEntidad As String
Private Const ENTIDAD_CUENTA As String = "C"
Private Const ENTIDAD_CCOSTO As String = "O"
Private Const COLORHABILITADO As Variant = &HC0E0FF
Private Const COLORDESABILITADO As Variant = &H80000005
'[Repetir en frmTCpr.
Private Const DIFERENCIAMASIMPORTEMN As Byte = 4, _
              DIFERENCIAMASIMPORTEME As Byte = 12, _
              DIFERENCIAMASCUENTA As Byte = 20, _
              DIFERENCIAMASCCOSTO As Byte = 28
Private Const CUENTASCONCCOSTO As Byte = 8
']
'[Repetir en frmTCprGrd y frmTCpr.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
']


Private Sub Form_Load()
   With frmTCprGrd.uorstCOCprDocCta
      frmTCprGrd.usConnStrgWher_COCprDocCta = " WHERE COCprDocCta.CodAux='" & frmTCpr.txtLlave(0).Text & "' AND COCprDocCta.CodTDc='" & frmTCpr.txtLlave(1).Text & "' AND COCprDocCta.SerDoc='" & frmTCpr.txtLlave(2).Text & "' AND COCprDocCta.NroDoc='" & frmTCpr.txtLlave(3).Text & "' AND COCprDocCta.TpoCnc=" & unIndice & " "
      If .State = adStateOpen Then .Close
      .Source = frmTCprGrd.usConnStrgSele_COCprDocCta & frmTCprGrd.usConnStrgWher_COCprDocCta & frmTCprGrd.usConnStrgOrde_COCprDocCta
      .Open
      .Properties("Unique Table").Value = "COCprDocCta"
   End With

   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = frmTCprGrd.uorstCOCprDocCta
 '[Propio del formulario.
   If unIndice <= CUENTASCONCCOSTO Then
      dgrCCosto.MarqueeStyle = dbgHighlightRow
      ppCCosto
   Else
      dgrCCosto.Visible = False
      Me.Height = 2650
   End If
 ']
End Sub

Private Sub Form_Activate()
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   upDatosGrid
 '[Propio del formulario.
   If unIndice <= CUENTASCONCCOSTO Then
      upDatosGrid_CCosto
   End If
 ']
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
'[ARREGLAR.
'   gpTUg_Resize Me
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   cmdSalir.Left = Me.Width - 820
   fraBuscar.Width = cmdSalir.Left - fraBuscar.Left - 50
   txtBuscar.Width = fraBuscar.Width - 240
'   dgrMain.Height = Me.ScaleHeight - 30 - picOpciones.Height '- uctEstado.Height
']ARREGLAR.
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   If unIndice <= CUENTASCONCCOSTO Then
'      frmTCprGrd.uorstCOCprDocCCo.Close
'      Set uorstCOCprDocCCo = Nothing
'   End If
End Sub

Public Sub cmdNuevo_Click()
Dim dvRegistro As Variant
   On Error GoTo Err
  
   Select Case psEntidad
   Case ENTIDAD_CUENTA
      gpTUg_Nuevo Me, frmTCprMasCta    'Cambiar Formulario de Datos.
   Case ENTIDAD_CCOSTO
      If frmTCprGrd.uorstCOCprDocCta.RecordCount = 0 Then
         MsgBox "No hay Cuentas creadas.", vbCritical
         Exit Sub
      Else
         With frmTCprGrd.uorstCOCta
            .MoveFirst
            .Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "'"
            If .EOF Then
               MsgBox "Esta Cuenta No Requiere Centro de Costos"
               Exit Sub
            ElseIf !IndCCo = INDCCO_INA Then
                   MsgBox "Esta Cuenta No Requiere Centro de Costos"
                   Exit Sub
            End If
         End With
      End If
      
      gpTUg_Nuevo Me, frmTCprMasCCo    'Cambiar Formulario de Datos.
      frmTCprGrd.uorstCOCprDocCCo.Requery
      dgrCCosto.SetFocus
'''      If frmTCprGrd.uorstCOCprDocCCo.RecordCount > 0 Then
'''         dvRegistro = frmTCprGrd.uorstCOCprDocCCo.Bookmark
'''         dgrCCosto.Refresh
'''         frmTCprGrd.uorstCOCprDocCCo.Bookmark = dvRegistro
'''         frmTCprGrd.uorstCOCprDocCCo.Find "CodCCo='" & frmTCprGrd.uorstCOCCo!CodCCo & "'"
'''         If frmTCprGrd.uorstCOCprDocCCo.EOF Then
'''            MsgBox "Error"
'''         End If
'''      End If
   End Select

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err

   Dim dvRegistroActual As Variant
   Dim dvCriterio As String

   Select Case psEntidad
   Case ENTIDAD_CUENTA
      If frmTCprGrd.uorstCOCprDocCta.RecordCount = 0 Then
         MsgBox "No hay Cuentas creadas.", vbCritical
         Exit Sub
      End If

      dvRegistroActual = frmTCprGrd.uorstCOCprDocCta.Bookmark

      With frmTCprMasCta               'Cambiar Formulario de Datos.
         .zbNuevo = False
         .upDatosDesconectados 1
       '[Deshabilitación de Llaves.    'Cambiar.
'         .txtLlave(0).Enabled = False
       ']
         .Caption = TEXT_MODIF & " " & Me.Caption

         .Show vbModal
      End With
      dgrMain.SetFocus

   Case ENTIDAD_CCOSTO
      If frmTCprGrd.uorstCOCprDocCCo.RecordCount = 0 Then
         MsgBox "No hay Centros de Costo creados.", vbCritical
         Exit Sub
      End If

      dvRegistroActual = frmTCprGrd.uorstCOCprDocCCo.Bookmark

      With frmTCprMasCCo               'Cambiar Formulario de Datos.
         .zbNuevo = False
         .upDatosDesconectados 1
       '[Deshabilitación de Llaves.    'Cambiar.
'         .txtLlave(0).Enabled = False
       ']
         .Caption = TEXT_MODIF & " " & Me.Caption

         .Show vbModal
      End With
      dgrCCosto.SetFocus
   End Select

   Exit Sub
Err:
   gpErrores
End Sub

Public Sub cmdEliminar_Click()
Dim dvRegistro As Variant
Dim dcCodCCo As String
   On Error GoTo Err

   Select Case psEntidad
   Case ENTIDAD_CUENTA
      If frmTCprGrd.uorstCOCprDocCta.RecordCount = 0 Then
         MsgBox "No hay Cuentas creadas.", vbCritical
         Exit Sub
      End If

      'Mensaje de verificación         'Cambiar.
      If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
'         frmTCprGrd.uocnnMain.BeginTrans
         frmTCprGrd.uorstCOCprDocCta.Delete
'         frmTCprGrd.uocnnMain.CommitTrans
      End If
      dgrMain.SetFocus

   Case ENTIDAD_CCOSTO
      If frmTCprGrd.uorstCOCprDocCCo.RecordCount = 0 Then
         MsgBox "No hay Centros de Costo creados.", vbCritical
         Exit Sub
      End If
      'Mensaje de verificación         'Cambiar.
      If MsgBox(TEXT_1021 & " " & Trim(dgrCCosto.Columns(0)) & " (" & Trim(dgrCCosto.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
'         frmTCprGrd.uocnnMain.BeginTrans
'         dcCodCCo = dgrCCosto.Columns(0)
'         dvRegistro = frmTCprGrd.uorstCOCprDocCCo.Bookmark
'         frmTCprGrd.uorstCOCprDocCCo.Requery
'         frmTCprGrd.uorstCOCprDocCCo.Bookmark = dvRegistro
'''         frmTCprGrd.uorstCOCprDocCCo.Find "cLlave2='" & frmTCpr.txtLlave(0).Text & frmTCpr.txtLlave(1).Text & frmTCpr.txtLlave(2).Text & frmTCpr.txtLlave(3).Text & frmTCprMasGrd.unIndice & frmTCprMasGrd.dgrMain.Columns(0).Text & dcCodCCo & "'"
'''         MsgBox (frmTCprGrd.uorstCOCprDocCCo!cLLave2)
''         frmTCprGrd.uorstCOCprDocCCo.Find "CodCCo='" & dcCodCCo & "'"
'         If Not frmTCprGrd.uorstCOCprDocCCo.EOF Then
            frmTCprGrd.uorstCOCprDocCCo.Delete
'         End If
''         frmTCprGrd.uocnnMain.CommitTrans
      End If
      dgrCCosto.SetFocus
   End Select

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Public Sub cmdRefrescar_Click()
   frmTCprGrd.uorstCOCprDocCta.Requery
   upDatosGrid
 '[Propio del formulario.
   frmTCprGrd.uorstCOCprDocCCo.Requery
   upDatosGrid_CCosto
 ']

   dgrMain.SetFocus
End Sub

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresión.  'Cambiar.
'   frmLPrd.Caption = "Listado de " & Me.Caption
'   frmLPrd.Show vbModal
 ']
End Sub

Private Sub cmdSalir_Click()
   Dim dnTotalCuentaMN, dnTotalCuentaME As Currency
   Dim dnTotalCCosto As Currency
   Dim dnValor As Byte
   
   With frmTCprGrd.uorstCOCprDocCta
      frmTCpr.TxtDato(unIndice + DIFERENCIAMASCUENTA).Text = ""
      If unIndice <= CUENTASCONCCOSTO Then
         frmTCpr.TxtDato(unIndice + DIFERENCIAMASCCOSTO).Text = ""
      End If
      
      If .RecordCount = 0 Then
         Call frmTCpr.upHabilitaCuenta(True, unIndice)
      Else
         Call frmTCpr.upHabilitaCuenta(False, unIndice)

        'Validación de Importes.
         .MoveFirst
         Do
            If unIndice <= CUENTASCONCCOSTO Then
               With frmTCprGrd.uorstCOCprDocCCo
                  If .RecordCount <> 0 Then
                     .MoveFirst
                     dnTotalCCosto = 0
                     Do
                        dnTotalCCosto = dnTotalCCosto + CDec(Choose(frmTCpr.unVerMonNac + 1, !ImpCCo_MN, !ImpCCo_ME))
                        .MoveNext
                     Loop Until .EOF
                     If CDec(Choose(frmTCpr.unVerMonNac + 1, frmTCprGrd.uorstCOCprDocCta!ImpCta_MN, frmTCprGrd.uorstCOCprDocCta!ImpCta_ME)) <> CDec(dnTotalCCosto) Then
                        .MoveFirst
                        MsgBox "El importe total de los Centros de Costo de la Cuenta " & frmTCprGrd.uorstCOCprDocCta!CodCta & " es " & dnTotalCCosto & " y debe ser " & Choose(frmTCpr.unVerMonNac + 1, frmTCprGrd.uorstCOCprDocCta!ImpCta_MN, frmTCprGrd.uorstCOCprDocCta!ImpCta_ME), vbCritical
                        dgrCCosto.SetFocus
                        Exit Sub
                     End If
                  End If
               End With
            End If
            dnTotalCuentaMN = dnTotalCuentaMN + !ImpCta_MN
            dnTotalCuentaME = dnTotalCuentaME + !ImpCta_ME
            .MoveNext
         Loop Until .EOF
         If (CDec(frmTCpr.TxtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Or (CDec(frmTCpr.TxtDato(unIndice + DIFERENCIAMASIMPORTEME).Text) <> CDec(dnTotalCuentaME)) Then
            .MoveFirst
            MsgBox "El importe total de las Cuentas Moneda Nacional : " & Format(dnTotalCuentaMN, FORMATO_NUM_1) & " y debe ser " & frmTCpr.TxtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text & Chr(13) & Space(48) & "Moneda Extranjera : " & Format(dnTotalCuentaME, FORMATO_NUM_1) & " y debe ser " & frmTCpr.TxtDato(unIndice + DIFERENCIAMASIMPORTEME).Text, vbCritical
            Exit Sub
         End If
      
        'Pintado de primera Cuenta y primer Centro de Costo.
         .MoveFirst
         frmTCpr.TxtDato(unIndice + DIFERENCIAMASCUENTA).Text = !CodCta
         If unIndice < CUENTASCONCCOSTO Then
            With frmTCprGrd.uorstCOCprDocCCo
               If .RecordCount <> 0 Then
                  .MoveFirst
                  frmTCpr.TxtDato(unIndice + DIFERENCIAMASCCOSTO).Text = !CodCCo
               End If
            End With
         End If
      End If
   End With

   Call frmTCpr.upActualizaMas(frmTCprMasGrd.unIndice, IIf(frmTCprGrd.uorstCOCprDocCta.RecordCount > 0, INDMASCTA_MAS, INDMASCTA_INI))

   Unload Me
End Sub

Private Sub dgrMain_GotFocus()
   psEntidad = ENTIDAD_CUENTA
   dgrMain.BackColor = COLORHABILITADO
   dgrCCosto.BackColor = COLORDESABILITADO
   dgrMain.HeadFont.Bold = True
   dgrCCosto.HeadFont.Bold = False
   With frmTCprGrd.uorstCOCta
      .MoveFirst
      If (Not .EOF) And frmTCprGrd.uorstCOCprDocCta.RecordCount > 0 Then
         .Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "'"
         If Not .EOF Then TxtDetaGrd(0).Text = !DetCta
      End If
   End With
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
''   pnColumnaOrd = ColIndex
''   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
''   txtBuscar = ""
''
''   frmTCprGrd.usConnStrgOrde = "ORDER BY "
'''   Select Case pnColumnaOrd            'Cambiar.
'''   Case 1, 2, 3
'''      psConnStrgOrde = psConnStrgOrde & "1, 2, 3"
'''   Case Else
''      frmTCprGrd.usConnStrgOrde = frmTCprGrd.usConnStrgOrde & pnColumnaOrd + 1
'''   End Select
''   With frmTCprGrd.uorstCOCprDocCta
''      .Close
''      .Source = frmTCprGrd.usConnStrgSele & frmTCprGrd.usConnStrgOrde
''      .Open
''   End With
''   Set dgrMain.DataSource = frmTCprGrd.uorstCOCprDocCta
''   upDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   If frmTCprGrd.uorstCOCprDocCta.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      frmTCprGrd.uorstCOCprDocCta.MoveFirst
   Case vbKeyEnd
      frmTCprGrd.uorstCOCprDocCta.MoveLast
   End Select
End Sub

Private Sub dgrMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   With frmTCprGrd.uorstCOCta
      .MoveFirst
      If Not .EOF Then
         If Not frmTCprGrd.uorstCOCprDocCta.EOF And Not frmTCprGrd.uorstCOCprDocCta.BOF Then
            If frmTCprGrd.uorstCOCprDocCta.Bookmark <> LastRow Then
               .Find "CodCta='" & Trim(dgrMain.Columns(0)) & "'"
               If Not .EOF Then TxtDetaGrd(0).Text = !DetCta
            End If
         End If
      End If
   End With
   ppCCosto
End Sub

Private Sub dgrCCosto_GotFocus()
   psEntidad = ENTIDAD_CCOSTO
   dgrMain.BackColor = COLORDESABILITADO
   dgrCCosto.BackColor = COLORHABILITADO
   dgrMain.HeadFont.Bold = False
   dgrCCosto.HeadFont.Bold = True
   With frmTCprGrd.uorstCOCCo
      .MoveFirst
      If (Not .EOF) And frmTCprGrd.uorstCOCprDocCCo.RecordCount > 0 Then
         .Find "CodCCo='" & frmTCprGrd.uorstCOCprDocCCo!CodCCo & "'"
         If Not .EOF Then TxtDetaGrd(1).Text = !DetCCo
      End If
   End With
End Sub

Private Sub dgrCCosto_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
''   pnColumnaOrd = ColIndex
''   fraBuscar.Caption = TEXT_BUSCA & dgrCCosto.Columns(pnColumnaOrd).Caption
''   txtBuscar = ""
''
''   frmTCprGrd.usConnStrgOrde = "ORDER BY "
'''   Select Case pnColumnaOrd            'Cambiar.
'''   Case 0
'''      frmTCprGrd.usConnStrgOrde = frmTCprGrd.usConnStrgOrde & "1"
'''   Case Else
''      frmTCprGrd.usConnStrgOrde = frmTCprGrd.usConnStrgOrde & pnColumnaOrd + 1
'''   End Select
''   With frmTCprGrd.uorstCOCprDocCCo
''      .Close
''      .Source = frmTCprGrd.usConnStrgSele_COCprDocCCo & frmTCprGrd.usConnStrgWher_COCprDocCCo & frmTCprGrd.usConnStrgOrde_COCprDocCCo
''      .Open
''   End With
''   Set dgrCCosto.DataSource = frmTCprGrd.uorstCOCprDocCCo
''   upDatosGrid_CCosto

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub dgrCCosto_KeyUp(KeyCode As Integer, Shift As Integer)
   If frmTCprGrd.uorstCOCprDocCCo.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      frmTCprGrd.uorstCOCprDocCCo.MoveFirst
   Case vbKeyEnd
      frmTCprGrd.uorstCOCprDocCCo.MoveLast
   End Select
End Sub

Private Sub dgrCCosto_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   With frmTCprGrd.uorstCOCCo
      .MoveFirst
      If (Not .EOF) Then
         If Not frmTCprGrd.uorstCOCprDocCCo.EOF And Not frmTCprGrd.uorstCOCprDocCCo.BOF Then
            If frmTCprGrd.uorstCOCprDocCCo.Bookmark <> LastRow Then
'               .Find "CodCCo='" & frmTCprGrd.uorstCOCprDocCCo!CodCCo & "'"
               .Find "CodCCo='" & Trim(dgrCCosto.Columns(0)) & "'"
               If Not .EOF Then TxtDetaGrd(1).Text = !DetCCo
            End If
         End If
      End If
   End With
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   Select Case psEntidad
   Case ENTIDAD_CUENTA
      With frmTCprGrd.uorstCOCprDocCta
         dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
         Select Case VarType(.Fields(pnColumnaOrd))
         Case vbString
            dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
         Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
            dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'        Case vbDate
'            dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
         End Select
         .Find dsCriterio, , , 1
         If .EOF = True Then
            .Bookmark = dvRegistroActual
         End If
      End With
']ARREGLAR.
   
   Case ENTIDAD_CCOSTO
      With frmTCprGrd.uorstCOCprDocCCo
         dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
         Select Case VarType(.Fields(pnColumnaOrd))
         Case vbString
            dsCriterio = dgrCCosto.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
         Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
            dsCriterio = dgrCCosto.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'        Case vbDate
'            dsCriterio = dgrCCosto.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
         End Select
         .Find dsCriterio, , , 1
         If .EOF = True Then
            .Bookmark = dvRegistroActual
         End If
      End With
']ARREGLAR.
   End Select
   
   Exit Sub
Err:
   If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
      frmTCprGrd.uorstCOCCo.Bookmark = dvRegistroActual
   Else
      gpErrores
   End If
End Sub

Public Sub upDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = "Cuenta"
            .Item(dnNum).Width = 850
''         Case 1
''            .Item(dnNum).Caption = "Descripción"
''            .Item(dnNum).Width = 3200
         Case 1
            .Item(dnNum).Caption = "Importe " & TPOMON_NAC_TXT_0
            .Item(dnNum).Width = 1150
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 2
            .Item(dnNum).Caption = "Importe " & TPOMON_EXT_TXT_0
            .Item(dnNum).Width = 1150
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[Código propio del formulario.

Public Sub ppCCosto()
   On Error GoTo Err
   
   With frmTCprGrd.uorstCOCprDocCCo
      If .State = adStateOpen Then .Close
      If frmTCprGrd.uorstCOCprDocCta.RecordCount <> 0 Then
''         psConnStrgWher_COCprDocCCo = "WHERE NroAut_Doc=" & frmTCpr.uorstCOCprDocCta!NroAut_Doc & " AND TpoCnc=" & unIndice & " AND CodCta='" & uorstCOCprDocCta!CodCta & "' "
         'frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE Concat(COCprDocCCo.CodAux, COCprDocCCo.SerDoc, COCprDocCCo.NroDoc)=" & frmTCprGrd.uorstMain!CodAux & frmTCprGrd.uorstMain!SerDoc & frmTCprGrd.uorstMain!NroDoc & " AND COCprDocCCo.TpoCnc=" & unIndice & " AND CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' And COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' And COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' And COCprDocCCo.TpoCnc='" & unIndice & "' And COCprDocCCo.CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "' "
      Else
''         psConnStrgWher_COCprDocCCo = "WHERE NroAut_Doc=" & frmTCpr.uorstCOCprDocCta!NroAut_Doc & " AND TpoCnc=" & unIndice & " AND CodCta='' "
         'frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE Concat(COCprDocCCo.CodAux, COCprDocCCo.SerDoc, COCprDocCCo.NroDoc, COCprDocCCo.TpoCnc, COCprDocCCo.CodCta)=" & frmTCprGrd.uorstMain!CodAux & frmTCprGrd.uorstMain!SerDoc & frmTCprGrd.uorstMain!NroDoc & " AND COCprDocCCo.TpoCnc=" & unIndice & " AND CodCta='' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' And COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' And COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' And COCprDocCCo.TpoCnc='" & unIndice & "' And COCprDocCCo.CodCta='' "
      End If
      .Source = frmTCprGrd.usConnStrgSele_COCprDocCCo & frmTCprGrd.usConnStrgWher_COCprDocCCo & frmTCprGrd.usConnStrgOrde_COCprDocCCo
      .Open
      .Properties("Unique Table").Value = "COCprDocCCo"
   End With
   Set dgrCCosto.DataSource = frmTCprGrd.uorstCOCprDocCCo
   upDatosGrid_CCosto
   Exit Sub
Err:
  If Err.Number = 3021 Then   'Se produce al llegar a EOF.
  Else
      gpErrores
  End If
End Sub

Public Sub upDatosGrid_CCosto()        'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrCCosto.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = "C.Costo"
            .Item(dnNum).Width = 850
'         Case 1
'            .Item(dnNum).Caption = "Descripción"
'            .Item(dnNum).Width = 3200
         Case 1
            .Item(dnNum).Caption = "Importe " & TPOMON_NAC_TXT_0
            .Item(dnNum).Width = 1150
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 2
            .Item(dnNum).Caption = "Importe " & TPOMON_EXT_TXT_0
            .Item(dnNum).Width = 1150
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

']Código propio del formulario.

Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdNuevo.Enabled = taOpciones(0)
   cmdEliminar.Enabled = taOpciones(1)
   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property


