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
         Index           =   1
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
              DIFERENCIAMASIMPORTEME As Byte = 15, _
              DIFERENCIAMASCUENTA As Byte = 26, _
              DIFERENCIAMASCCOSTO As Byte = 37
Private Const CUENTASCONCCOSTO As Byte = 11
']
'[Repetir en frmTCprGrd y frmTCpr.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
']

Private Sub cmdImprimir_Click(Index As Integer)
 '[Datos del formulario de impresión.  'Cambiar.
'   frmLPrd.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
'   frmLPrd.Show vbModal
 ']
End Sub
Private Sub Form_Load()
  
  With frmTCprGrd.uorstCOCprDocCta
    frmTCprGrd.usConnStrgWher_COCprDocCta = "WHERE COCprDocCta.codemp='" & gsCodEmp & "' AND COCprDocCta.pdoano='" & gsAnoAct & "' "
    frmTCprGrd.usConnStrgWher_COCprDocCta = frmTCprGrd.usConnStrgWher_COCprDocCta & "AND COCprDocCta.CodAux='" & frmTCpr.txtLlave(0).Text & "' AND COCprDocCta.CodTDc='" & frmTCpr.txtLlave(1).Text & "' "
    frmTCprGrd.usConnStrgWher_COCprDocCta = frmTCprGrd.usConnStrgWher_COCprDocCta & "AND COCprDocCta.SerDoc='" & frmTCpr.txtLlave(2).Text & "' AND COCprDocCta.NroDoc='" & frmTCpr.txtLlave(3).Text & "' AND COCprDocCta.TpoCnc=" & unIndice & " "
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
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  Me.Caption = Choose(gsIdioma, "Detalle", "Detail")
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
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
         MsgBox Choose(gsIdioma, "No hay Cuentas creadas.", "There are not created accounts"), vbCritical
         Exit Sub
      Else
         With frmTCprGrd.uorstCOCta
            .MoveFirst
            .Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!codcta & "'"
            If .EOF Then
               MsgBox Choose(gsIdioma, "Esta Cuenta No Requiere Centro de Costos", "This account does not need Cost Center")
               Exit Sub
            ElseIf !IndCCo = INDCCO_INA Then
               MsgBox Choose(gsIdioma, "Esta Cuenta No Requiere Centro de Costos", "This account does not need Cost Center")
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
         MsgBox Choose(gsIdioma, "No hay Cuentas creadas.", "There are not created accounts"), vbCritical
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
         MsgBox Choose(gsIdioma, "No hay Centros de Costo creados.", "There are not created Cost Centers."), vbCritical
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
         MsgBox Choose(gsIdioma, "No hay Cuentas creadas.", "There are not created accounts."), vbCritical
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
         MsgBox Choose(gsIdioma, "No hay Centros de Costo creados.", "There are not created Cost Centers."), vbCritical
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


Private Sub cmdSalir_Click()
  Dim dnTotalCuentaMN As Currency, dnTotalCuentaME As Currency
  Dim dnTotalCCostoMN As Currency, dnTotalCCostoME As Currency
  Dim dsMensaje As String, dnImporte As Currency
  Dim dnValor As Byte
  
  With frmTCprGrd.uorstCOCprDocCta
    frmTCpr.txtDato(unIndice + DIFERENCIAMASCUENTA).Text = ""
    If unIndice <= CUENTASCONCCOSTO Then frmTCpr.txtDato(unIndice + DIFERENCIAMASCCOSTO).Text = ""
    
    If .RecordCount = 0 Then
      frmTCpr.upHabilitaCuenta True, unIndice
    Else
      frmTCpr.upHabilitaCuenta False, unIndice
      'Validación de Importes.
      .MoveFirst
      Do
        If unIndice <= CUENTASCONCCOSTO Then
          With frmTCprGrd.uorstCOCprDocCCo
            If .RecordCount <> 0 Then
              .MoveFirst
              dnTotalCCostoMN = 0
              dnTotalCCostoME = 0
              Do
                dnTotalCCostoMN = dnTotalCCostoMN + CDec(!impcco_mn)
                dnTotalCCostoME = dnTotalCCostoME + CDec(!impcco_me)
                .MoveNext
              Loop Until .EOF
              If ((CDec(frmTCprGrd.uorstCOCprDocCta!impcta_mn) <> CDec(dnTotalCCostoMN)) Or (CDec(frmTCprGrd.uorstCOCprDocCta!impcta_me) <> CDec(dnTotalCCostoME))) Then
                dsMensaje = Choose(gsIdioma, "Cuenta ", "Account ") & frmTCprGrd.uorstCOCprDocCta!codcta & Choose(gsIdioma, " El importe total de los Centros de Costo ", " The total amount of Cost Center ") & Chr(13)
                dsMensaje = dsMensaje & Choose(gsIdioma, "Moneda Nacional : ", "National Currency : ") & Format(dnTotalCCostoMN, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser : ", " and must be : ") & Format(frmTCprGrd.uorstCOCprDocCta!impcta_mn, FORMATO_NUM_1)
                dsMensaje = dsMensaje & Choose(gsIdioma, " La diferencia es de ", " The difference is ") & Trim(CStr(Abs(dnTotalCCostoMN - CDec(frmTCprGrd.uorstCOCprDocCta!impcta_mn)))) & Chr(13)
                dsMensaje = dsMensaje & Choose(gsIdioma, "Moneda Extranjera : ", "Foreign Currency : ") & Format(dnTotalCCostoME, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser : ", " and must be : ") & Format(frmTCprGrd.uorstCOCprDocCta!impcta_me, FORMATO_NUM_1)
                dsMensaje = dsMensaje & Choose(gsIdioma, " La diferencia es de ", " The difference is ") & Trim(CStr(Abs(dnTotalCCostoME - CDec(frmTCprGrd.uorstCOCprDocCta!impcta_me))))
                .MoveFirst
                If ((Abs(CDec(dnTotalCCostoMN) - CDec(frmTCprGrd.uorstCOCprDocCta!impcta_mn)) <= 0.02) And (Abs(CDec(dnTotalCCostoME) - CDec(frmTCprGrd.uorstCOCprDocCta!impcta_me)) <= 0.02)) Then
                  If MsgBox(dsMensaje & Chr(13) & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then dgrCCosto.SetFocus: Exit Sub
                  .MoveLast
                  ' Actualizo los importes de centro de costos
                  If (CDec(frmTCprGrd.uorstCOCprDocCta!impcta_mn) <> CDec(dnTotalCCostoMN)) Then
                    dnImporte = CDec(dnTotalCCostoMN) - CDec(frmTCprGrd.uorstCOCprDocCta!impcta_mn)
                    frmTCprGrd.uorstCOCprDocCCo!impcco_mn = CDec(frmTCprGrd.uorstCOCprDocCCo!impcco_mn) + (dnImporte * -1)
                  End If
                  If (CDec(frmTCprGrd.uorstCOCprDocCta!impcta_me) <> CDec(dnTotalCCostoME)) Then
                    dnImporte = CDec(dnTotalCCostoME) - CDec(frmTCprGrd.uorstCOCprDocCta!impcta_me)
                    frmTCprGrd.uorstCOCprDocCCo!impcco_me = CDec(frmTCprGrd.uorstCOCprDocCCo!impcco_me) + (dnImporte * -1)
                  End If
                  frmTCprGrd.uorstCOCprDocCCo!UsrMdf = gsAbvUsr
                  frmTCprGrd.uorstCOCprDocCCo!FyHMdf = Now
                  frmTCprGrd.uorstCOCprDocCCo.Update
                  .MoveFirst
                Else
                  MsgBox dsMensaje, vbCritical
                  dgrCCosto.SetFocus: Exit Sub
                End If
              End If
            End If
          End With
        End If
        dnTotalCuentaMN = dnTotalCuentaMN + !impcta_mn
        dnTotalCuentaME = dnTotalCuentaME + !impcta_me
        .MoveNext
      Loop Until .EOF
      If (CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Or (CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text) <> CDec(dnTotalCuentaME)) Then
        dsMensaje = Choose(gsIdioma, "El importe total de las Cuentas ", "The total amount of Acccounts ") & Chr(13)
        dsMensaje = dsMensaje & Choose(gsIdioma, "Moneda Nacional : ", "National Currency : ") & Format(dnTotalCuentaMN, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser : ", " and must be : ") & frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text & Space(5)
        dsMensaje = dsMensaje & Choose(gsIdioma, " La diferencia es de ", " The difference is ") & Trim(CStr(Abs(dnTotalCuentaMN - CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text)))) & Chr(13)
        dsMensaje = dsMensaje & Choose(gsIdioma, "Moneda Extranjera : ", "Foreign Currency : ") & Format(dnTotalCuentaME, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser : ", " and must be : ") & frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text & Space(5)
        dsMensaje = dsMensaje & Choose(gsIdioma, " La diferencia es de ", " The difference is ") & Trim(CStr(Abs(dnTotalCuentaME - CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text)))) & Chr(13)
        .MoveFirst
        If ((Abs(CDec(dnTotalCuentaMN) - CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text)) <= 0.02) And (Abs(CDec(dnTotalCuentaME) - CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text)) <= 0.02)) Then
          If MsgBox(dsMensaje & Chr(13) & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then dgrCCosto.SetFocus: Exit Sub
          .MoveLast
          ' Actualizo los importes de cuentas
          If (CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Then
            dnImporte = CDec(dnTotalCuentaMN) - CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text)
            frmTCprGrd.uorstCOCprDocCta!impcta_mn = CDec(frmTCprGrd.uorstCOCprDocCta!impcta_mn) + (dnImporte * -1)
          End If
          If (CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text) <> CDec(dnTotalCuentaME)) Then
            dnImporte = CDec(dnTotalCuentaME) - CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text)
            frmTCprGrd.uorstCOCprDocCta!impcta_me = CDec(frmTCprGrd.uorstCOCprDocCta!impcta_me) + (dnImporte * -1)
          End If
          frmTCprGrd.uorstCOCprDocCta!UsrMdf = gsAbvUsr
          frmTCprGrd.uorstCOCprDocCta!FyHMdf = Now
          frmTCprGrd.uorstCOCprDocCta.Update
          
          ' Actualizo los importes de centro de costo
          If frmTCprGrd.uorstCOCprDocCCo.RecordCount > 0 Then
            frmTCprGrd.uorstCOCprDocCCo.MoveLast
            If (CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Then
              frmTCprGrd.uorstCOCprDocCCo!impcco_mn = CDec(frmTCprGrd.uorstCOCprDocCCo!impcco_mn) + (dnImporte * -1)
            End If
            If (CDec(frmTCpr.txtDato(unIndice + DIFERENCIAMASIMPORTEME).Text) <> CDec(dnTotalCuentaME)) Then
              frmTCprGrd.uorstCOCprDocCCo!impcco_me = CDec(frmTCprGrd.uorstCOCprDocCCo!impcco_me) + (dnImporte * -1)
            End If
            frmTCprGrd.uorstCOCprDocCCo!UsrMdf = gsAbvUsr
            frmTCprGrd.uorstCOCprDocCCo!FyHMdf = Now
            frmTCprGrd.uorstCOCprDocCCo.Update
            .MoveFirst
            frmTCprGrd.uorstCOCprDocCCo.MoveFirst
          End If
        Else
          MsgBox dsMensaje, vbCritical
          dgrCCosto.SetFocus: Exit Sub
        End If
      End If
      
      'Pintado de primera Cuenta y primer Centro de Costo.
      .MoveFirst
      frmTCpr.txtDato(unIndice + DIFERENCIAMASCUENTA).Text = !codcta
      If unIndice < CUENTASCONCCOSTO Then
        With frmTCprGrd.uorstCOCprDocCCo
          If .RecordCount <> 0 Then
            .MoveFirst
            frmTCpr.txtDato(unIndice + DIFERENCIAMASCCOSTO).Text = !codcco
          End If
        End With
      End If
    End If
  End With
  frmTCpr.upActualizaMas frmTCprMasGrd.unIndice, IIf(frmTCprGrd.uorstCOCprDocCta.RecordCount > 0, INDMASCTA_MAS, INDMASCTA_INI)
  
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
         .Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!codcta & "'"
         If Not .EOF Then TxtDetaGrd(0).Text = IIf(IsNull(!detcta), "", !detcta)
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
               If Not .EOF Then TxtDetaGrd(0).Text = IIf(IsNull(!detcta), "", !detcta)
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
   With frmTCprGrd.uorstCoCCo
      .MoveFirst
      If (Not .EOF) And frmTCprGrd.uorstCOCprDocCCo.RecordCount > 0 Then
         .Find "CodCCo='" & frmTCprGrd.uorstCOCprDocCCo!codcco & "'"
         If Not .EOF Then TxtDetaGrd(1).Text = IIf(IsNull(!DetCCo), "", !DetCCo)
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
   With frmTCprGrd.uorstCoCCo
      .MoveFirst
      If (Not .EOF) Then
         If Not frmTCprGrd.uorstCOCprDocCCo.EOF And Not frmTCprGrd.uorstCOCprDocCCo.BOF Then
            If frmTCprGrd.uorstCOCprDocCCo.Bookmark <> LastRow Then
'               .Find "CodCCo='" & frmTCprGrd.uorstCOCprDocCCo!CodCCo & "'"
               .Find "CodCCo='" & Trim(dgrCCosto.Columns(0)) & "'"
               If Not .EOF Then TxtDetaGrd(1).Text = IIf(IsNull(!DetCCo), "", !DetCCo)
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
      frmTCprGrd.uorstCoCCo.Bookmark = dvRegistroActual
   Else
      gpErrores
   End If
End Sub

Public Sub upDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
            
  dgrMain.Caption = Choose(gsIdioma, "Cuenta", "Account")
  With dgrMain.Columns
     For dnNum = 0 To .Count - 1
        Select Case dnNum
        Case 0
           .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
           .Item(dnNum).Width = 850
        Case 1
           .Item(dnNum).Caption = Choose(gsIdioma, "Importe ", "Amount ") & TPOMON_NAC_TXT_0
           .Item(dnNum).Width = 1100
           .Item(dnNum).Alignment = dbgRight
           .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
        Case 2
           .Item(dnNum).Caption = Choose(gsIdioma, "Importe ", "Amount ") & TPOMON_EXT_TXT_0
           .Item(dnNum).Width = 1100
           .Item(dnNum).Alignment = dbgRight
           .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
        Case 3
           .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
           .Item(dnNum).Width = 2300
        Case 4
           .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
           .Item(dnNum).Width = 1200
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
         frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.codemp='" & frmTCprGrd.uorstMain!codemp & "' AND COCprDocCCo.pdoano='" & frmTCprGrd.uorstMain!pdoano & "' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' AND COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' AND COCprDocCCo.TpoCnc='" & unIndice & "' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.Orden='" & frmTCprGrd.uorstCOCprDocCta!orden & "' AND COCprDocCCo.CodCta='" & frmTCprGrd.uorstCOCprDocCta!codcta & "' "
      Else
         frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.codemp='" & frmTCprGrd.uorstMain!codemp & "' AND COCprDocCCo.pdoano='" & frmTCprGrd.uorstMain!pdoano & "' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' AND COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' "
         frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' And COCprDocCCo.TpoCnc='" & unIndice & "' AND COCprDocCCo.Orden='" & frmTCprGrd.uorstCOCprDocCta!orden & "' AND COCprDocCCo.CodCta='' "
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
         
  dgrCCosto.Caption = Choose(gsIdioma, "Centro de Costo", "Cost Center")
  With dgrCCosto.Columns
     For dnNum = 0 To .Count - 1
        Select Case dnNum
        Case 0
           .Item(dnNum).Caption = Choose(gsIdioma, "C.Costo", "Cost Center")
           .Item(dnNum).Width = 950
'         Case 1
'            .Item(dnNum).Caption = "Descripción"
'            .Item(dnNum).Width = 3200
        Case 1
           .Item(dnNum).Caption = Choose(gsIdioma, "Importe ", "Amount ") & TPOMON_NAC_TXT_0
           .Item(dnNum).Width = 1150
           .Item(dnNum).Alignment = dbgRight
           .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
        Case 2
           .Item(dnNum).Caption = Choose(gsIdioma, "Importe ", "Amount ") & TPOMON_EXT_TXT_0
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
   cmdImprimir(1).Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property


