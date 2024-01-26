VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmTRteVtaMasGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4890
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
         Picture         =   "frmTRteVtaMasGrd.frx":0000
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
         Picture         =   "frmTRteVtaMasGrd.frx":014A
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
         Picture         =   "frmTRteVtaMasGrd.frx":0294
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
         Picture         =   "frmTRteVtaMasGrd.frx":0396
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
         Picture         =   "frmTRteVtaMasGrd.frx":0498
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
         Picture         =   "frmTRteVtaMasGrd.frx":059A
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
Attribute VB_Name = "frmTRteVtaMasGrd"
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
'[Repetir en frmTrteVta.
Private Const DIFERENCIAMASIMPORTEMN As Byte = 7, _
              DIFERENCIAMASCUENTA As Byte = 14, _
              DIFERENCIAMASCCOSTO As Byte = 21
Private Const CUENTASCONCCOSTO As Byte = 7
']
'[Repetir en frmTRteVtaGrd y fmrTVta.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
']

Private Sub dgrCCosto_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  With frmTRteVtaGrd.uorstCoCCo
    .MoveFirst
    If (Not .EOF) Then
      If Not frmTRteVtaGrd.uorstCoRteVtaCCo.EOF And Not frmTRteVtaGrd.uorstCoRteVtaCCo.BOF Then
        If frmTRteVtaGrd.uorstCoRteVtaCCo.Bookmark <> LastRow Then
          .Find "codcco='" & Trim(dgrCCosto.Columns(0)) & "'"
          If Not .EOF Then TxtDetaGrd(1).Text = IIf(IsNull(!detcco), "", !detcco)
        End If
      End If
    End If
  End With
End Sub

Private Sub Form_Load()
  With frmTRteVtaGrd.uorstCoRteVtaCta
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta = "WHERE CoRteVtaCta.codemp='" & gsCodEmp & "' "
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta & "AND CoRteVtaCta.sernegocio='" & frmTRteVta.txtLlave(0).Text & "' "
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta & "AND CoRteVtaCta.nronegocio='" & frmTRteVta.txtLlave(1).Text & "' AND CoRteVtaCta.TpoCnc=" & unIndice & " "
    If .State = adStateOpen Then .Close
    .Source = frmTRteVtaGrd.usConnStrgSele_CoRteVtaCta & frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta & frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCta
    .Open
    .Properties("Unique Table").Value = "CoRteVtaCta"
  End With
  
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = frmTRteVtaGrd.uorstCoRteVtaCta
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
'      frmTRteVtaGrd.uorstCOrteVtaCCo.Close
'      Set uorstCOVtaDocCCo = Nothing
'   End If
End Sub

Public Sub cmdNuevo_Click()
  On Error GoTo Err
  
  Select Case psEntidad
   Case ENTIDAD_CUENTA
    gpTUg_Nuevo Me, frmTRteVtaMasCta    'Cambiar Formulario de Datos.
   Case ENTIDAD_CCOSTO
    If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay Cuentas creadas.", "There are not created accounts"), vbCritical
      Exit Sub
    Else
      With frmTRteVtaGrd.uorstCOCta
        .MoveFirst
        .Find "CodCta='" & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & "'"
        If .EOF Then
          MsgBox Choose(gsIdioma, "Esta Cuenta No Requiere Centro de Costos", "This account does not need Cost Center")
          Exit Sub
        ElseIf !indcco = INDCCO_INA Then
          MsgBox Choose(gsIdioma, "Esta Cuenta No Requiere Centro de Costos", "This account does not need Cost Center")
          Exit Sub
        End If
      End With
    End If
    
    gpTUg_Nuevo Me, frmTRteVtaMasCCo    'Cambiar Formulario de Datos.
    dgrCCosto.SetFocus
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
    If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay Cuentas creadas.", "There are not created accounts"), vbCritical
      Exit Sub
    End If
    dvRegistroActual = frmTRteVtaGrd.uorstCoRteVtaCta.Bookmark
    
    With frmTRteVtaMasCta               'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
      .Caption = TEXT_MODIF & " " & Me.Caption
      .Show vbModal
    End With
    dgrMain.SetFocus
    
   Case ENTIDAD_CCOSTO
    If frmTRteVtaGrd.uorstCoRteVtaCCo.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay Centros de Costo creados.", "There are not created Cost Center"), vbCritical
      Exit Sub
    End If
    dvRegistroActual = frmTRteVtaGrd.uorstCoRteVtaCCo.Bookmark
    With frmTRteVtaMasCCo               'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
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
  On Error GoTo Err
  
  Select Case psEntidad
   Case ENTIDAD_CUENTA
    If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay Cuentas creadas.", "There are not created accounts"), vbCritical
      Exit Sub
    End If
    
    'Mensaje de verificación         'Cambiar.
    If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      frmTRteVtaGrd.uorstCoRteVtaCta.Delete
    End If
    dgrMain.SetFocus
   Case ENTIDAD_CCOSTO
    If frmTRteVtaGrd.uorstCoRteVtaCCo.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay Centros de Costo creados.", "There are not created Cost Center"), vbCritical
      Exit Sub
    End If
    
    'Mensaje de verificación         'Cambiar.
    If MsgBox(TEXT_1021 & " " & Trim(dgrCCosto.Columns(0)) & " (" & Trim(dgrCCosto.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      frmTRteVtaGrd.uorstCoRteVtaCCo.Delete
    End If
    dgrCCosto.SetFocus
  End Select
  
  Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Public Sub cmdRefrescar_Click()
  frmTRteVtaGrd.uorstCoRteVtaCta.Requery
  upDatosGrid
  '[Propio del formulario.
  frmTRteVtaGrd.uorstCoRteVtaCCo.Requery
  upDatosGrid_CCosto
  ']
  dgrMain.SetFocus
End Sub

Public Sub cmdImprimir_Click(Index As Integer)
 '[Datos del formulario de impresión.  'Cambiar.
'   frmLPrd.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
'   frmLPrd.Show vbModal
 ']
End Sub

Private Sub cmdSalir_Click()
  Dim dnTotalCuentaMN As Currency, dnTotalCuentaME As Currency
  Dim dnTotalCCostoMN As Currency, dnTotalCCostoME As Currency
  Dim dsMensaje As String, dnImporte As Currency
  Dim dnValor As Byte
     
  With frmTRteVtaGrd.uorstCoRteVtaCta
    frmTRteVta.txtDato(unIndice + DIFERENCIAMASCUENTA).Text = ""
    If unIndice <= CUENTASCONCCOSTO Then frmTRteVta.txtDato(unIndice + DIFERENCIAMASCCOSTO).Text = ""
    
    If .RecordCount = 0 Then
      frmTRteVta.upHabilitaCuenta True, unIndice
    Else
      frmTRteVta.upHabilitaCuenta False, unIndice
      'Validación de Importes.
      .MoveFirst
      Do
        If unIndice <= CUENTASCONCCOSTO Then
          With frmTRteVtaGrd.uorstCoRteVtaCCo
            If .RecordCount <> 0 Then
              .MoveFirst
              dnTotalCCostoMN = 0
              Do
                dnTotalCCostoMN = dnTotalCCostoMN + CDec(!porimpcco)
                .MoveNext
              Loop Until .EOF
              
              If (CDec(dnTotalCCostoMN) <> 100) Then
                dsMensaje = Choose(gsIdioma, "Cuenta ", "Account ") & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & Choose(gsIdioma, " El porcentaje total de los Centros de Costo ", " The total percentage of Cost Center ") & Chr(13)
                dsMensaje = dsMensaje & frmTRteVta.cboTpoMon.Text & " : " & Format(dnTotalCCostoMN, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser : ", " and must be : ") & Format(100, FORMATO_NUM_1)
                dsMensaje = dsMensaje & Choose(gsIdioma, " La diferencia es de ", " The difference is ") & Trim(CStr(Abs(dnTotalCCostoMN - 100))) & Chr(13)
                .MoveFirst
                If (Abs(CDec(dnTotalCCostoMN) - 100) <= 0.02) Then
                  If MsgBox(dsMensaje & Chr(13) & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then dgrCCosto.SetFocus: Exit Sub
                  .MoveLast
                  ' Actualizo los importes de centro de costos
                  If (CDec(dnTotalCCostoMN) <> 100) Then
                    dnImporte = CDec(dnTotalCCostoMN) - 100
                    frmTRteVtaGrd.uorstCoRteVtaCCo!porimpcco = CDec(frmTRteVtaGrd.uorstCoRteVtaCCo!porimpcco) + (dnImporte * -1)
                  End If
                  frmTRteVtaGrd.uorstCoRteVtaCCo!UsrMdf = gsAbvUsr
                  frmTRteVtaGrd.uorstCoRteVtaCCo!FyHMdf = Now
                  frmTRteVtaGrd.uorstCoRteVtaCCo.Update
                  .MoveFirst
                Else
                  MsgBox dsMensaje, vbCritical
                  dgrCCosto.SetFocus: Exit Sub
                End If
              End If
            End If
          End With
        End If
        dnTotalCuentaMN = dnTotalCuentaMN + !porimpcta
        .MoveNext
      Loop Until .EOF
      dnTotalCuentaME = Round(dnTotalCuentaMN / 100, 2)
      dnTotalCuentaMN = Round(CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) * dnTotalCuentaME, 2)
      
      If (CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Then
        dsMensaje = Choose(gsIdioma, "El importe total de las Cuentas ", "The total amount of Acccounts ") & Chr(13)
        dsMensaje = dsMensaje & frmTRteVta.cboTpoMon.Text & " : " & Format(dnTotalCuentaMN, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser : ", " and must be : ") & frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text & Space(5)
        dsMensaje = dsMensaje & Choose(gsIdioma, " La diferencia es de ", " The difference is ") & Trim(CStr(Abs(dnTotalCuentaMN - CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text)))) & Chr(13)
        .MoveFirst
        If ((Abs(CDec(dnTotalCuentaMN) - CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text)) <= 0.02)) Then
          If MsgBox(dsMensaje & Chr(13) & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then dgrCCosto.SetFocus: Exit Sub
          .MoveLast
          ' Actualizo los importes de cuentas
          If (CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Then
            dnImporte = CDec(dnTotalCuentaMN) - CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text)
            frmTRteVtaGrd.uorstCoRteVtaCta!porimpcta = CDec(frmTRteVtaGrd.uorstCoRteVtaCta!porimpcta) + (dnImporte * -1)
          End If
          frmTRteVtaGrd.uorstCoRteVtaCta!UsrMdf = gsAbvUsr
          frmTRteVtaGrd.uorstCoRteVtaCta!FyHMdf = Now
          frmTRteVtaGrd.uorstCoRteVtaCta.Update
          
          ' Actualizo los importes de centro de costo
          If frmTRteVtaGrd.uorstCoRteVtaCCo.RecordCount > 0 Then
            frmTRteVtaGrd.uorstCoRteVtaCCo.MoveLast
            If (CDec(frmTRteVta.txtDato(unIndice + DIFERENCIAMASIMPORTEMN).Text) <> CDec(dnTotalCuentaMN)) Then
              frmTRteVtaGrd.uorstCoRteVtaCCo!porimpcco = CDec(frmTRteVtaGrd.uorstCoRteVtaCCo!porimpcco) + (dnImporte * -1)
            End If
            frmTRteVtaGrd.uorstCoRteVtaCCo!UsrMdf = gsAbvUsr
            frmTRteVtaGrd.uorstCoRteVtaCCo!FyHMdf = Now
            frmTRteVtaGrd.uorstCoRteVtaCCo.Update
            .MoveFirst
            frmTRteVtaGrd.uorstCoRteVtaCCo.MoveFirst
          End If
        Else
          MsgBox dsMensaje, vbCritical
          dgrCCosto.SetFocus: Exit Sub
        End If
      End If
      
      'Pintado de primera Cuenta y primer Centro de Costo.
      .MoveFirst
      frmTRteVta.txtDato(unIndice + DIFERENCIAMASCUENTA).Text = !codcta
      If unIndice < CUENTASCONCCOSTO Then
        With frmTRteVtaGrd.uorstCoRteVtaCCo
          If .RecordCount <> 0 Then
            .MoveFirst
            frmTRteVta.txtDato(unIndice + DIFERENCIAMASCCOSTO).Text = !codcco
          End If
        End With
      End If
    End If
  End With
  
  Unload Me
End Sub

Private Sub dgrMain_GotFocus()
   psEntidad = ENTIDAD_CUENTA
   dgrMain.BackColor = COLORHABILITADO
   dgrCCosto.BackColor = COLORDESABILITADO
   dgrMain.HeadFont.Bold = True
   dgrCCosto.HeadFont.Bold = False
   With frmTRteVtaGrd.uorstCOCta
      .MoveFirst
      If (Not .EOF) And frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount > 0 Then
         .Find "CodCta='" & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & "'"
         If Not .EOF Then TxtDetaGrd(0).Text = IIf(IsNull(!detcta), "", !detcta)
      End If
   End With
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
'   pnColumnaOrd = ColIndex
'   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
'   txtBuscar = ""
'
'   frmTRteVtaGrd.usConnStrgOrde_COrteVtaCta = "ORDER BY "
''   Select Case pnColumnaOrd            'Cambiar.
''   Case 1, 2, 3
''      psConnStrgOrde = psConnStrgOrde & "1, 2, 3"
''   Case Else
'      frmTRteVtaGrd.usConnStrgOrde_COrteVtaCta = frmTRteVtaGrd.usConnStrgOrde_COrteVtaCta & pnColumnaOrd + 1
''   End Select
'   With frmTRteVtaGrd.uorstCoRteVtaCta
'      .Close
'      .Source = frmTRteVtaGrd.usConnStrgSele_COrteVtaCta & frmTRteVtaGrd.usConnStrgOrde_COrteVtaCta
'      .Open
'   End With
'   Set dgrMain.DataSource = frmTRteVtaGrd.uorstCoRteVtaCta
'   upDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
  If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount = 0 Then Exit Sub
  
  Select Case KeyCode
   Case vbKeyHome
    frmTRteVtaGrd.uorstCoRteVtaCta.MoveFirst
   Case vbKeyEnd
    frmTRteVtaGrd.uorstCoRteVtaCta.MoveLast
  End Select
End Sub

Private Sub dgrMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  With frmTRteVtaGrd.uorstCOCta
    .MoveFirst
    If Not .EOF Then
      If Not frmTRteVtaGrd.uorstCoRteVtaCta.EOF And Not frmTRteVtaGrd.uorstCoRteVtaCta.BOF Then
        If frmTRteVtaGrd.uorstCoRteVtaCta.Bookmark <> LastRow Then
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
  With frmTRteVtaGrd.uorstCoCCo
    .MoveFirst
    If (Not .EOF) And frmTRteVtaGrd.uorstCoRteVtaCCo.RecordCount > 0 Then
      .Find "CodCCo='" & frmTRteVtaGrd.uorstCoRteVtaCCo!codcco & "'"
      If Not .EOF Then TxtDetaGrd(1).Text = IIf(IsNull(!detcco), "", !detcco)
    End If
  End With
End Sub

Private Sub dgrCCosto_HeadClick(ByVal ColIndex As Integer)
  On Error GoTo Err
  
  pnColumnaOrd = ColIndex
  fraBuscar.Caption = TEXT_BUSCA & dgrCCosto.Columns(pnColumnaOrd).Caption
  txtBuscar = ""
  
  frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo = "ORDER BY "
  frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo & pnColumnaOrd + 1
  With frmTRteVtaGrd.uorstCoRteVtaCCo
    .Close
    .Source = frmTRteVtaGrd.usConnStrgSele_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo
    .Open
  End With
  Set dgrCCosto.DataSource = frmTRteVtaGrd.uorstCoRteVtaCCo
  upDatosGrid_CCosto
  
  Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub dgrCCosto_KeyUp(KeyCode As Integer, Shift As Integer)
  If frmTRteVtaGrd.uorstCoRteVtaCCo.RecordCount = 0 Then Exit Sub
  Select Case KeyCode
   Case vbKeyHome
    frmTRteVtaGrd.uorstCoRteVtaCCo.MoveFirst
   Case vbKeyEnd
    frmTRteVtaGrd.uorstCoRteVtaCCo.MoveLast
  End Select
End Sub

Private Sub txtBuscar_Change()
  On Error GoTo Err
  
  Dim dsCriterio As String
  Dim dvRegistroActual As Variant
  
  Select Case psEntidad
   Case ENTIDAD_CUENTA
    With frmTRteVtaGrd.uorstCoRteVtaCta
      dvRegistroActual = .Bookmark
      '[ARREGLAR: Búsqueda con distintos tipos de columna.
      Select Case VarType(.Fields(pnColumnaOrd))
       Case vbString
        dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
       Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
        dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
      End Select
      .Find dsCriterio, , , 1
      If .EOF = True Then
        .Bookmark = dvRegistroActual
      End If
    End With
    ']ARREGLAR.
    
   Case ENTIDAD_CCOSTO
    With frmTRteVtaGrd.uorstCoRteVtaCCo
      dvRegistroActual = .Bookmark
      
      '[ARREGLAR: Búsqueda con distintos tipos de columna.
      Select Case VarType(.Fields(pnColumnaOrd))
       Case vbString
        dsCriterio = dgrCCosto.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
       Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
        dsCriterio = dgrCCosto.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
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
    frmTRteVtaGrd.uorstCoRteVtaCta.Bookmark = dvRegistroActual
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
        .Item(dnNum).Caption = Choose(gsIdioma, "Porcentaje", "percentage")
        .Item(dnNum).Width = 1100
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
        .Item(dnNum).Width = 3400
       Case 3
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
  With frmTRteVtaGrd.uorstCoRteVtaCCo
    If .State = adStateOpen Then .Close
    If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount <> 0 Then
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = "WHERE CoRteVtaCCo.codemp ='" & frmTRteVtaGrd.uorstMain!codemp & "' "
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND CoRteVtaCCo.sernegocio='" & frmTRteVtaGrd.uorstMain!sernegocio & "' "
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND CoRteVtaCCo.nronegocio='" & frmTRteVtaGrd.uorstMain!nronegocio & "' AND CoRteVtaCCo.TpoCnc=" & unIndice & " "
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND CoRteVtaCCo.Orden='" & frmTRteVtaGrd.uorstCoRteVtaCta!orden & "' AND CoRteVtaCCo.CodCta='" & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & "' "
    Else
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = "WHERE CoRteVtaCCo.codemp='" & frmTRteVtaGrd.uorstMain!codemp & "' "
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND CoRteVtaCCo.sernegocio='" & frmTRteVtaGrd.uorstMain!sernegocio & "' "
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND CoRteVtaCCo.nronegocio='" & frmTRteVtaGrd.uorstMain!nronegocio & "' AND CoRteVtaCCo.TpoCnc=" & unIndice & " "
      frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND CoRteVtaCCo.Orden='" & frmTRteVtaGrd.uorstCoRteVtaCta!orden & "' AND CoRteVtaCCo.CodCta='' "
    End If
    .Source = frmTRteVtaGrd.usConnStrgSele_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo
    .Open
    .Properties("Unique Table").Value = "CoRteVtaCCo"
  End With
  Set dgrCCosto.DataSource = frmTRteVtaGrd.uorstCoRteVtaCCo
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
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "Porcentaje", "percentage")
        .Item(dnNum).Width = 1150
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case Else
        .Item(dnNum).Visible = False
      End Select
    Next dnNum
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

