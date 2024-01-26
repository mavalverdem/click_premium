VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMCieGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   4395
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   1635
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8475
      _ExtentX        =   14949
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
      Caption         =   "[Título 1]"
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
      ScaleWidth      =   8475
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   8475
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
         Picture         =   "frmMCieGrd.frx":0000
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
         Left            =   7750
         Picture         =   "frmMCieGrd.frx":014A
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
         Picture         =   "frmMCieGrd.frx":0294
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
         Picture         =   "frmMCieGrd.frx":0396
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
         Picture         =   "frmMCieGrd.frx":0498
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
         Picture         =   "frmMCieGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   2115
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3731
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
      Caption         =   "[Título 2]"
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
Attribute VB_Name = "frmMCieGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain_0 As ADODB.Recordset
Public uorstMain_1 As ADODB.Recordset
Public uorstUltiItem As ADODB.Recordset
Private psConnStrgSele_0 As String, _
        psConnStrgOrde_0 As String
Private psConnStrgSele_1 As String, _
        psConnStrgWher_1 As String, _
        psConnStrgOrde_1 As String
Private pnColumnaOrd As Integer
Private pnEntidad As Integer
Private Const ENTIDAD_0 As Integer = 0, _
              ENTIDAD_1 As Integer = 1
Private Const MENSAJE_ENTIDAD_0 As String = "Asientos de Cierre creados", _
              MENSAJE_ENTIDAD_1 As String = "Cuentas creadas"
Private Const COLORHABILITADO   As Variant = &HC0E0FF, _
              COLORDESABILITADO As Variant = &H80000005

'[Propio del formulario.
Public uorstCOCta As ADODB.Recordset
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   psConnStrgSele_0 = "SELECT NroCie, DetCie, " _
                    & "  UsrCre, FyHCre, UsrMdf, FyHMdf " _
                    & "FROM COCIE "
   psConnStrgOrde_1 = "ORDER BY 1"
   psConnStrgSele_1 = "SELECT COCIECTA.CodCta, b.DetCta, COCIECTA.NroCie, COCIECTA.NroIte, " _
                    & "  COCIECTA.IndHTr, COCIECTA.TpoHTr, COCIECTA.TpoHT1, COCIECTA.FmlCie, " _
                    & "  COCIECTA.TpoCtb, COCIECTA.IndCCt, COCIECTA.TpoCtbI, COCIECTA.ImpMNI, COCIECTA.IndAMo, " _
                    & "  COCIECTA.UsrCre, COCIECTA.FyHCre, COCIECTA.UsrMdf, COCIECTA.FyHMdf " _
                    & "FROM COCIECTA " _
                    & "  LEFT JOIN COCta b ON COCIECTA.CodCta=b.CodCta "
   psConnStrgOrde_1 = "ORDER BY 3, 4"

   Set uocnnMain = New ADODB.Connection
   Set uorstMain_0 = New ADODB.Recordset
   Set uorstMain_1 = New ADODB.Recordset
   Set uorstCOCta = New ADODB.Recordset
   Set uorstUltiItem = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With uorstMain_0
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele_0 & psConnStrgOrde_0
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "COCIE"
   End With
   With uorstMain_1
      .ActiveConnection = uocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
   End With
   With uorstCOCta
      .ActiveConnection = uocnnMain
      .Source = "SELECT a.CodCta, a.DetCta, a.IndMoe " _
              & "FROM COCta a"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstUltiItem
      .ActiveConnection = uocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
'      .Open
   End With
 ']
   
   dgrMain(0).Caption = "Asiento de Cierre" 'Cambiar.
   dgrMain(1).Caption = "Cuenta"        'Cambiar.
   dgrMain(0).MarqueeStyle = dbgHighlightRow
   dgrMain(1).MarqueeStyle = dbgHighlightRow
   Set dgrMain(0).DataSource = uorstMain_0
   ppStrg_1
End Sub

Private Sub Form_Activate()
   'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   upDatosGrid 0
   fraBuscar.Caption = TEXT_BUSCA & dgrMain(0).Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'[ARREGLAR. Definir el procedimiento a seguir.
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
']ARREGLAR.
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
'''  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
'''   cmdSalir.Left = Me.Width - 820
'''   fraBuscar.Width = cmdSalir.Left - fraBuscar.Left - 50
'''   txtBuscar.Width = fraBuscar.Width - 240
''''   dgrMain(0).Height = Me.ScaleHeight - 30 - picOpciones.Height '- uctEstado.Height
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Cambiar Recordsets.
   If uorstCOCta.State = adStateOpen Then uorstCOCta.Close
   If uorstMain_1.State = adStateOpen Then uorstMain_1.Close
   If uorstMain_0.State = adStateOpen Then uorstMain_0.Close
   If uocnnMain.State = adStateOpen Then uocnnMain.Close
   Set uorstCOCta = Nothing
   Set uorstMain_1 = Nothing
   Set uorstMain_0 = Nothing
   Set uocnnMain = Nothing
End Sub

Public Sub cmdNuevo_Click()
   Select Case pnEntidad
   Case ENTIDAD_0
'      gpTUg_Nuevo Me, frmMCie       'Cambiar Formulario de Datos.
      With frmMCie
         .zbNuevo = True   'Tiene que ir primero para que el load lo coja evaluado.
         .Caption = TEXT_NUEVO & " " & frmMCieGrd.Caption
         .upDatosPredeterminados
      
         .Show vbModal
      End With
   
      frmMCieGrd.dgrMain(0).SetFocus
   
   Case ENTIDAD_1
      If uorstMain_0.RecordCount = 0 Then
         MsgBox "No hay " & MENSAJE_ENTIDAD_1 & ".", vbCritical
         Exit Sub
      End If
      
'      gpTUg_Nuevo Me, frmMCieLin       'Cambiar Formulario de Datos.
      With frmMCieCta
         .zbNuevo = True   'Tiene que ir primero para que el load lo coja evaluado.
         .Caption = TEXT_NUEVO & " " & frmMCieGrd.Caption
         .upDatosPredeterminados
      
         .Show vbModal
      End With
   
      frmMCieGrd.dgrMain(1).SetFocus
   End Select
End Sub

Private Sub cmdRevisar_click()
   On Error GoTo Err
   
   Select Case pnEntidad
   Case ENTIDAD_0
      If uorstMain_0.RecordCount = 0 Then
         MsgBox "No hay " & MENSAJE_ENTIDAD_0 & ".", vbCritical
         Exit Sub
      End If
   
      With frmMCie                  'Cambiar Formulario de Datos.
         .zbNuevo = False
         .upDatosDesconectados 1
       '[Deshabilitación de Llaves.    'Cambiar.
         .txtLlave(0).Enabled = False
       ']
         .Caption = TEXT_MODIF & " " & Me.Caption
         
         .Show vbModal
      End With
   
      dgrMain(0).SetFocus
   
   Case ENTIDAD_1
      If uorstMain_1.RecordCount = 0 Then
         MsgBox "No hay " & MENSAJE_ENTIDAD_1 & ".", vbCritical
         Exit Sub
      End If
   
      With frmMCieCta                  'Cambiar Formulario de Datos.
         .zbNuevo = False
         .upDatosDesconectados 1
       '[Deshabilitación de Llaves.    'Cambiar.
         .txtLlave(0).Enabled = False
       ']
         .Caption = TEXT_MODIF & " " & Me.Caption
         
         .Show vbModal
      End With
   
      dgrMain(1).SetFocus
   End Select
  
   Exit Sub
Err:
   gpErrores
End Sub

Public Sub cmdEliminar_Click()
   On Error GoTo Err
  
   Select Case pnEntidad
   Case ENTIDAD_0
      If uorstMain_0.RecordCount = 0 Then
         MsgBox "No hay " & MENSAJE_ENTIDAD_0 & ".", vbCritical
         Exit Sub
      End If
   
      'Mensaje de verificación            'Cambiar.
      If MsgBox(TEXT_1021 & " " & Trim(dgrMain(0).Columns(0)) & " (" & Trim(dgrMain(0).Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
         uocnnMain.BeginTrans
         uorstMain_0.Delete
         uocnnMain.CommitTrans
         
         ppStrg_1
      End If
      dgrMain(0).SetFocus
   
   Case ENTIDAD_1
      If uorstMain_1.RecordCount = 0 Then
         MsgBox "No hay " & MENSAJE_ENTIDAD_0 & ".", vbCritical
         Exit Sub
      End If
   
      'Mensaje de verificación            'Cambiar.
      If MsgBox(TEXT_1021 & " " & Trim(dgrMain(1).Columns(0)) & " (" & Trim(dgrMain(1).Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
         uocnnMain.BeginTrans
         uorstMain_1.Delete adAffectCurrent
         uocnnMain.CommitTrans
      
      End If
      dgrMain(1).SetFocus
   End Select

   Exit Sub
Err:
   gpErrores
   
   uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
   uorstMain_0.Requery
   uorstMain_1.Requery
   upDatosGrid 0
   upDatosGrid 1

   dgrMain(0).SetFocus
End Sub

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresión.  'Cambiar.
   frmLEFi.Caption = "Listado de " & Me.Caption
   frmLEFi.Show vbModal
 ']
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_GotFocus(Index As Integer)
   Select Case Index
   Case ENTIDAD_0
      pnEntidad = ENTIDAD_0
      dgrMain(0).BackColor = COLORHABILITADO
      dgrMain(1).BackColor = COLORDESABILITADO
      dgrMain(0).HeadFont.Bold = True
      dgrMain(1).HeadFont.Bold = False
   Case ENTIDAD_1
      pnEntidad = ENTIDAD_1
      dgrMain(0).BackColor = COLORDESABILITADO
      dgrMain(1).BackColor = COLORHABILITADO
      dgrMain(0).HeadFont.Bold = False
      dgrMain(1).HeadFont.Bold = True
   End Select
End Sub

Private Sub dgrMain_HeadClick(Index As Integer, ByVal ColIndex As Integer)
   Select Case Index
   Case ENTIDAD_0
      pnColumnaOrd = ColIndex
      fraBuscar.Caption = TEXT_BUSCA & dgrMain(0).Columns(pnColumnaOrd).Caption
      txtBuscar = ""
   
      psConnStrgOrde_0 = "ORDER BY "
'      Select Case pnColumnaOrd            'Cambiar.
'      Case 1, 2, 3
'         psConnStrgOrde_0 = psConnStrgOrde_0 & "1, 2, 3"
'      Case Else
         psConnStrgOrde_0 = psConnStrgOrde_0 & pnColumnaOrd + 1
'      End Select
      With uorstMain_0
         .Close
         .Source = psConnStrgSele_0 & psConnStrgOrde_0
         .Open
      End With
      Set dgrMain(0).DataSource = uorstMain_0
      upDatosGrid 0
   Case ENTIDAD_1
      pnColumnaOrd = ColIndex
      fraBuscar.Caption = TEXT_BUSCA & dgrMain(1).Columns(pnColumnaOrd).Caption
      txtBuscar = ""
   
      psConnStrgOrde_1 = "ORDER BY "
'      Select Case pnColumnaOrd            'Cambiar.
'      Case 1, 2, 3
'         psConnStrgOrde_1 = psConnStrgOrde_1 & "1, 2, 3"
'      Case Else
         psConnStrgOrde_1 = psConnStrgOrde_1 & pnColumnaOrd + 1
'      End Select
      With uorstMain_1
         .Close
         .Source = psConnStrgSele_1 & psConnStrgOrde_1
         .Open
      End With
      Set dgrMain(1).DataSource = uorstMain_1
      upDatosGrid 1
   End Select

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case Index
   Case ENTIDAD_0
      Select Case KeyCode
      Case vbKeyHome
         uorstMain_0.MoveFirst
      Case vbKeyEnd
         uorstMain_0.MoveLast
      End Select
   Case ENTIDAD_1
      Select Case KeyCode
      Case vbKeyHome
         uorstMain_1.MoveFirst
      Case vbKeyEnd
         uorstMain_1.MoveLast
      End Select
   End Select
End Sub

Private Sub dgrMain_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
   Select Case Index
   Case ENTIDAD_0
      ppStrg_1
   End Select
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   Select Case pnEntidad
   Case ENTIDAD_0
      With uorstMain_0
         dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
         Select Case VarType(.Fields(pnColumnaOrd))
         Case vbString
            dsCriterio = dgrMain(0).Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
         Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
            dsCriterio = dgrMain(0).Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'        Case vbDate
'            dsCriterio = dgrMain(0).Columns(pnColumnaOrd).DataField & " = " & txtBuscar
         End Select
         .Find dsCriterio, , , 1
         If .EOF = True Then
            .Bookmark = dvRegistroActual
         End If
      End With
']ARREGLAR.
   
   Case ENTIDAD_1
      With uorstMain_1
         dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
         Select Case VarType(.Fields(pnColumnaOrd))
         Case vbString
            dsCriterio = dgrMain(1).Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
         Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
            dsCriterio = dgrMain(1).Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'        Case vbDate
'            dsCriterio = dgrMain(1).Columns(pnColumnaOrd).DataField & " = " & txtBuscar
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
   If Err.Number = 3001 Then   'Se produce al llegar a EOF del recordset.
      Select Case pnEntidad
      Case ENTIDAD_0
         uorstMain_0.Bookmark = dvRegistroActual
      Case ENTIDAD_1
         uorstMain_1.Bookmark = dvRegistroActual
      End Select
   Else
      MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
   End If
End Sub

Public Sub upDatosGrid(tnIndex As Integer) 'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   Select Case tnIndex
   Case ENTIDAD_0
      With dgrMain(0).Columns
         For dnNum = 0 To .Count - 1
            Select Case dnNum
            Case 0
               .Item(dnNum).Caption = "Nro."
               .Item(dnNum).Width = 700
            Case 1
               .Item(dnNum).Caption = "Descripción"
               .Item(dnNum).Width = 5000
            Case Else
               .Item(dnNum).Visible = False
            End Select
         Next
      End With
   Case ENTIDAD_1
      With dgrMain(1).Columns
         For dnNum = 0 To .Count - 1
            Select Case dnNum
            Case 0
               .Item(dnNum).Caption = "Cuenta"
               .Item(dnNum).Width = 1300
            Case 1
               .Item(dnNum).Caption = "Descripción"
               .Item(dnNum).Width = 5180
            Case Else
               .Item(dnNum).Visible = False
            End Select
         Next
      End With
   End Select
End Sub

'[Código propio del formulario.

Public Sub ppStrg_1()
   On Error GoTo Err
   
   With uorstMain_1
      If .State = adStateOpen Then .Close
      If uorstMain_0.Status = adRecDBDeleted Then
         uorstMain_0.MoveNext
         If uorstMain_0.EOF Then uorstMain_0.MovePrevious
      End If
      psConnStrgWher_1 = "WHERE NroCie='" & uorstMain_0!NroCie & "' "
      .Source = psConnStrgSele_1 & psConnStrgWher_1 & psConnStrgOrde_1
      .Open
      .Properties("Unique Table").Value = "COCIECTA"
   End With
   Set dgrMain(1).DataSource = uorstMain_1
   upDatosGrid 1
   
   Exit Sub
Err:
  If Err.Number = 3021 Then   'Se produce al llegar a EOF.
  Else
     MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  End If
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

