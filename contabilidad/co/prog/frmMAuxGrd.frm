VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMAuxGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   9075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdOtros 
      Cancel          =   -1  'True
      Caption         =   "&Onp/Afp"
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
      Left            =   3600
      Picture         =   "frmMAuxGrd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   720
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   6165
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
      ScaleWidth      =   9075
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   9075
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
         Picture         =   "frmMAuxGrd.frx":0188
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
         Left            =   4365
         TabIndex        =   0
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   200
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdSalir 
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
         Picture         =   "frmMAuxGrd.frx":02D2
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmMAuxGrd.frx":041C
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
         Picture         =   "frmMAuxGrd.frx":051E
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
         Picture         =   "frmMAuxGrd.frx":0620
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
         Picture         =   "frmMAuxGrd.frx":0722
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMAuxGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer

'[Propio del formulario.
'Public uorstTGDtt As ADODB.Recordset
Public uorstMai2 As ADODB.Recordset
Public psConnStrgSel2 As String, _
       psConnStrgOrd2 As String, _
       psConnStrgCon2 As String
']
'Public uocnnNoGrabable As ADODB.Connection
Public uorstCoEntidadPen As ADODB.Recordset '2014-05-01 RR.HH afecto afp/onp


Private Sub cmdOtros_Click()
    frmMOnpGrd.Show vbModal
End Sub

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   psConnStrgSele = "SELECT CodAux, RazAux, RucAux, DirAux, Email, rubro, "
   psConnStrgSele = psConnStrgSele & "codemp, UsrCre, FyHCre, UsrMdf, FyHMdf, EstAux, "
   psConnStrgSele = psConnStrgSele & "IndCli, IndPrv, IndOtr, TpoPer, TpoDci "
   psConnStrgSele = psConnStrgSele & "FROM TgAux "
   psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
   psConnStrgOrde = "ORDER BY 1"
   
   psConnStrgSel2 = "SELECT CodAux, NomAux, ApePatAux, ApeMatAux, codtdi, numdci, "
   psConnStrgSel2 = psConnStrgSel2 & "codemp, UsrCre, FyHCre, UsrMdf, FyHMdf "
   psConnStrgSel2 = psConnStrgSel2 & "FROM tgauxnat "
   psConnStrgSel2 = psConnStrgSel2 & "WHERE codemp='" & gsCodEmp & "' "
   psConnStrgCon2 = "AND codaux="
   psConnStrgOrd2 = " ORDER BY 1"

   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   Set uorstMai2 = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "TgAux"
'      .Properties("Unique Table").Value = "a"
   End With
   With uorstMai2
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSel2 & psConnStrgOrd2
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
'      .Properties("Unique Table").Value = "TgAux"
   End With
 ']
'ini 2014-08-05 RR.HH afecto afp/onp
  Set uorstCoEntidadPen = New ADODB.Recordset
  With uorstCoEntidadPen
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.Codafp, a.Desafp "
    .Source = .Source & "FROM Coentidadpen a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    '.Source = .Source & "AND a.IndCli=1 "
    .Source = .Source & "AND a.Estadoafp='" & ESTAUX_ACT & "'"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With

'  With uorstCoEntidadPen
'    .ActiveConnection = uocnnNoGrabable
'    '.ActiveConnection = uocnnMain
'    .Source = "SELECT a.Codafp, a.Desafp "
'    .Source = .Source & "FROM Coentidadpen a "
'    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
'    '.Source = .Source & "AND a.IndCli=1 "
'    .Source = .Source & "AND a.Estadoafp='" & ESTAUX_ACT & "'"
'    '     .CursorLocation = adUseClient   'Es el Default.
'    .CursorType = adOpenDynamic
'    .LockType = adLockReadOnly
'    .Open
'  End With

'fin 2014-08-05 RR.HH afecto afp/onp

  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  ']
   
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain
End Sub

Private Sub Form_Activate()
   'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   ppDatosGrid
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   gpTUg_Resize2 Me
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
'   uorstMai2.Close
   uorstMain.Close
   uorstCoEntidadPen.Close '2014-05-01 RR.HH afecto afp/onp
   uocnnMain.Close
   Set uorstMai2 = Nothing
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
   
   Set uorstCoEntidadPen = Nothing '2014-05-01 RR.HH afecto afp/onp
End Sub

Public Sub cmdNuevo_Click()
    frmMAux.puorstOnp_Insert '2014-05-01 RR.HH afecto afp/onp

   gpTUg_Nuevo Me, frmMAux             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
   
   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

'[  Data de Persona Natural
   With uorstMai2
      .Close
      .Source = psConnStrgSel2 & psConnStrgCon2 & "'" & Trim(dgrMain.Columns(0)) & "'" & psConnStrgOrd2
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
']
   With frmMAux                        'Cambiar Formulario de Datos.
    .zbNuevo = False
    .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
    .txtLlave(0).Enabled = False
    ']
    
    .puorstOnp_Insert '2014-05-01 RR.HH afecto afp/onp
    
    .Caption = TEXT_MODIF & " " & Me.Caption
    .Show vbModal
   End With
   
   dgrMain.SetFocus
  
   Exit Sub
Err:
   gpErrores
End Sub

Public Sub cmdEliminar_Click()
Dim ppBorra As Integer

   On Error GoTo Err

   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If
   ppBorra = 0

   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
'[  Tabla de Personas Naturales (Tabla hijo)
      With uorstMai2
         .Close
         .Source = psConnStrgSel2 & psConnStrgCon2 & "'" & Trim(dgrMain.Columns(0)) & "'" & psConnStrgOrd2
         .Open
      End With
      If uorstMai2.RecordCount = 0 Then
         ppBorra = 1
      End If
'ini 2014-08-05 RR.HH afecto afp/onp
     frmMAuxGrd.uocnnMain.Execute "DELETE FROM codonpafp " & _
    "WHERE codemp='" & gsCodEmp & "' AND Codaux='" & Trim(dgrMain.Columns(0)) & "'"
    '"' AND Codafp='" & .Fields("CodAfp") & "'"
'fin 2014-08-05 RR.HH afecto afp/onp
']
      uocnnMain.BeginTrans
      If ppBorra = 0 Then
         uorstMai2.Delete
      End If
      uorstMain.Delete
      uocnnMain.CommitTrans
   End If
   dgrMain.SetFocus

   Exit Sub
Err:
   gpErrores
   
   uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
   gpTUg_Refrescar Me
End Sub

Public Sub cmdImprimir_Click(Index As Integer)
  If uorstMain.RecordCount = 0 Then
     MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
     Exit Sub
  End If
 '[Datos del formulario de impresión.  'Cambiar.
   frmLAux.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLAux.Show vbModal
 ']
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = "ORDER BY "
'   Select Case pnColumnaOrd            'Cambiar.
'   Case 1, 2, 3
'      psConnStrgOrde = psConnStrgOrde & "1, 2, 3"
'   Case Else
      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
'   End Select
   With uorstMain
      .Close
      .Source = psConnStrgSele & psConnStrgOrde
      .Open
   End With
   Set dgrMain.DataSource = uorstMain
   ppDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   If uorstMain.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      uorstMain.MoveFirst
   Case vbKeyEnd
      uorstMain.MoveLast
   End Select
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With uorstMain
      dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
      Select Case VarType(.Fields(pnColumnaOrd))
      Case vbString
         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
      Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'     Case vbDate
'         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
      End Select
      .Find dsCriterio, , , 1
      If .EOF = True Then
         .Bookmark = dvRegistroActual
      End If
   End With
']ARREGLAR.
   
   Exit Sub
Err:
   If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
      uorstMain.Bookmark = dvRegistroActual
   Else
      gpErrores
   End If
End Sub

Public Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Código", "Code")
             .Item(dnNum).Width = 100 * (uorstMain.Fields("CodAux").DefinedSize)
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
             .Item(dnNum).Width = 100 * (uorstMain.Fields("RazAux").DefinedSize - 12)
             
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[Código propio del formulario.

']

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

