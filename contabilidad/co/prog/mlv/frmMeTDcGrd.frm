VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMeTDcGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   6390
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   8475
      _ExtentX        =   14949
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
      ScaleWidth      =   8475
      TabIndex        =   9
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
         Picture         =   "frmMeTDcGrd.frx":0000
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
         Picture         =   "frmMeTDcGrd.frx":014A
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
         Picture         =   "frmMeTDcGrd.frx":0294
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
         Picture         =   "frmMeTDcGrd.frx":0396
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
         Picture         =   "frmMeTDcGrd.frx":0498
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
         Picture         =   "frmMeTDcGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMeTDcGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As adodb.Connection
Public uorstMain As adodb.Recordset
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer

'[Propio del formulario.
'Public uorstTGDtt As ADODB.Recordset
']

Private Sub Form_Load()
  Dim s_Sentencia As String
  Dim porstTmp As adodb.Recordset
  Dim nRegistro As Long
  
 '[Recordsets                          'Cambiar.
  psConnStrgSele = "SELECT cntbkeyocur, destpodoc, abvtipdoc, codtpodoc, sgntpodoc"
  psConnStrgSele = psConnStrgSele & " FROM tmptpodoc "
  psConnStrgOrde = "ORDER BY 1, 4"
 
  Set uocnnMain = New adodb.Connection
  Set uorstMain = New adodb.Recordset
  With uocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = gfParaOracle
    .Open
  End With
   
  ' Creo la tabla si no existe
  s_Sentencia = "SELECT COUNT(*) AS nExiste FROM dba_tables WHERE table_name='" & UCase("tmptpodoc") & "'"
  Set porstTmp = New adodb.Recordset
  With porstTmp
    If .State = adStateOpen Then .Close
    .ActiveConnection = uocnnMain
    .Source = s_Sentencia
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
    nRegistro = !nExiste
    .Close
  End With
  Set porstTmp = Nothing
  If nRegistro = 0 Then
    s_Sentencia = "CREATE TABLE tmptpodoc("
    s_Sentencia = s_Sentencia & " cntbkeyocur char(17) NOT NULL,"
    s_Sentencia = s_Sentencia & " codtpodoc char(2),"
    s_Sentencia = s_Sentencia & " destpodoc char(35),"
    s_Sentencia = s_Sentencia & " abvtipdoc char(3),"
    s_Sentencia = s_Sentencia & " sgntpodoc char(1) DEFAULT '+',"
    s_Sentencia = s_Sentencia & " PRIMARY KEY (cntbkeyocur))"
    uocnnMain.Execute s_Sentencia, nRegistro
  End If
   
  ' Inserto registros no existentes
  s_Sentencia = "INSERT INTO tmptpodoc(cntbkeyocur, destpodoc, abvtipdoc)"
  s_Sentencia = s_Sentencia & " SELECT DISTINCT cntbkeyocur, LPAD(cntbdatocur, 25) AS destpodoc,"
  s_Sentencia = s_Sentencia & " LPAD(cntbkeyocur, 3) AS abvtipdoc"
  s_Sentencia = s_Sentencia & " FROM cndmtabl tmp"
  s_Sentencia = s_Sentencia & " WHERE cntbcodtabl=1"
  s_Sentencia = s_Sentencia & " AND NOT EXISTS(SELECT * FROM tmptpodoc dro"
  s_Sentencia = s_Sentencia & " WHERE tmp.cntbkeyocur=dro.cntbkeyocur)"
  s_Sentencia = s_Sentencia & " AND cntbkeyocur NOT IN('!!!', 'ANY', 'NBD', 'ODP', 'PLZ')"
  s_Sentencia = s_Sentencia & " ORDER BY cntbkeyocur"
  uocnnMain.Execute s_Sentencia
   
  With uorstMain
    .ActiveConnection = uocnnMain
    .Source = psConnStrgSele & psConnStrgOrde
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "tmptpodoc"
  End With
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
  
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   uorstMain.Close
   uocnnMain.Close
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Public Sub cmdNuevo_Click()
   gpTUg_Nuevo Me, frmMeTDc             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
   
   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   With frmMeTDc                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
      .txtLlave(0).Enabled = False
    ']
      .Caption = TEXT_MODIF & " " & Me.Caption
      
      .Show vbModal
   End With

   dgrMain.SetFocus
  
   Exit Sub
Err:
   gpErrores
End Sub

Public Sub cmdEliminar_Click()
   On Error GoTo Err

   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      uocnnMain.BeginTrans
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

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresión.  'Cambiar.
   frmLTDc.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLTDc.Show vbModal
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
            .Item(dnNum).Caption = "Código"
            .Item(dnNum).Width = 100 * (uorstMain.Fields!cntbkeyocur.DefinedSize - 8)
         Case 1
            .Item(dnNum).Caption = "Descripción"
            .Item(dnNum).Width = 100 * (uorstMain.Fields!destpodoc.DefinedSize)
         Case 2
            .Item(dnNum).Caption = "Abreviado"
            .Item(dnNum).Width = 100 * (uorstMain.Fields!abvtipdoc.DefinedSize + 5)
         Case 3
            .Item(dnNum).Caption = "Equivalencia"
            .Item(dnNum).Width = 100 * (uorstMain.Fields!codtpodoc.DefinedSize + 8)
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
   cmdNuevo.Enabled = False
   cmdEliminar.Enabled = taOpciones(1)
   cmdImprimir.Enabled = False
End Property

