VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMDetracGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
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
         Picture         =   "frmMDetracGrd.frx":0000
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
         Picture         =   "frmMDetracGrd.frx":014A
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
         Picture         =   "frmMDetracGrd.frx":0294
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
         Picture         =   "frmMDetracGrd.frx":0396
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
         Picture         =   "frmMDetracGrd.frx":0498
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
         Picture         =   "frmMDetracGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMDetracGrd"
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
Public uorstcodetrac As ADODB.Recordset
Public uorstCOLIB As ADODB.Recordset
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   psConnStrgSele = "SELECT coddetrac, "
   psConnStrgSele = psConnStrgSele & Choose(gsIdioma, "detdetrac, ", "detdetracx, ")
   psConnStrgSele = psConnStrgSele & Choose(gsIdioma, "detdetracx, ", "detdetrac, ")
   psConnStrgSele = psConnStrgSele & "(CASE estdetrac WHEN '" & ESTCCO_ACT & "' THEN '" & ESTCCO_ACT_TXT & "' ELSE '" & ESTCCO_INA_TXT & "' END) AS cestdetrac, "
   psConnStrgSele = psConnStrgSele & "pctdetrac, "
   psConnStrgSele = psConnStrgSele & "estdetrac, "
   psConnStrgSele = psConnStrgSele & "codemp, UsrCre, FyHCre, UsrMdf, FyHMdf "
   psConnStrgSele = psConnStrgSele & "FROM codetrac "
   psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
   'psConnStrgSele = psConnStrgSele & "AND pdoano='" & gsAnoAct & "'"
   psConnStrgOrde = "ORDER BY 1"
   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   Set uorstcodetrac = New ADODB.Recordset
   Set uorstCOLIB = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
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
      .Properties("Unique Table").Value = "codetrac"
'2015-06-26 proceso detrac       .Properties("Unique Table").Value = "CODRO"
'      .Properties("Unique Table").Value = "a"
   End With
   With uorstcodetrac
      .ActiveConnection = uocnnMain
      .Source = "SELECT a.coddetrac, a.detdetrac, a.detdetracx "
      .Source = .Source & "FROM codetrac a "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      '.Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstCOLIB
      .ActiveConnection = uocnnMain
      .Source = "SELECT a.CodLIB, a.DeSLIB "
      .Source = .Source & "FROM COLIB a "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "order by 1 asc "
      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With

 ']
      
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
   gpTUg_Resize Me
End Sub
Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   uorstMain.Close
   uocnnMain.Close
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub
Public Sub cmdNuevo_Click()
   gpTUg_Nuevo Me, frmMDetrac             'Cambiar Formulario de Datos.
End Sub
Public Sub cmdRevisar_click()
   On Error GoTo Err
   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If
   With frmMDetrac                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitaci�n de Llaves.       'Cambiar.
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
   '[No Pertenece al Formulario - Agregado por Angel
Dim dcCodEli As String
   ']
   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If
Dim xxSelect As String
Dim xxTabla As String


Dim xErroInteRefer As Integer
xErroInteRefer = 0
Dim uorstTmp1 As ADODB.Recordset
Set uorstTmp1 = New ADODB.Recordset
xxTabla = "cocprdoc"
xxSelect = ""
xxSelect = xxSelect & "select tsadetrac,pctdetrac from " & xxTabla & " a "
xxSelect = xxSelect & "where codemp='" & gsCodEmp & "' and pdoano='" & gsAnoAct & "' "
xxSelect = xxSelect & "and tsadetrac='" & Trim(dgrMain.Columns(0)) & "'"

Set uorstTmp1 = fRstOpen(uocnnMain, uorstTmp1, xxSelect, adOpenDynamic, adLockOptimistic)
If Not uorstTmp1.EOF Then
    xErroInteRefer = 1
End If
Set uorstTmp1 = New ADODB.Recordset
xxTabla = "covtadoc"
xxSelect = ""
xxSelect = xxSelect & "select tsadetrac,pctdetrac from " & xxTabla & " a "
xxSelect = xxSelect & "where codemp='" & gsCodEmp & "' and pdoano='" & gsAnoAct & "' "
xxSelect = xxSelect & "and tsadetrac='" & Trim(dgrMain.Columns(0)) & "'"

Set uorstTmp1 = fRstOpen(uocnnMain, uorstTmp1, xxSelect, adOpenDynamic, adLockOptimistic)
If Not uorstTmp1.EOF Then
    xErroInteRefer = 1
End If
fRstClose uorstTmp1
If xErroInteRefer = 1 Then MsgBox (TEXT_9022): Exit Sub
'   uorstMain.Close
   
'ini 2015-06-26 proceso detrac
'   '[No Pertenece al Formulario - Agregado por Angel
'   dcCodEli = frmMDetracGrd.uorstMain!coddetrac
'   If Len(Trim(dcCodEli)) = 2 Then
'      With uorstcodetrac
'         .Close
'         .Source = "SELECT a.coddetrac, a.detdetrac, a.detdetracx "
'         .Source = .Source & "FROM codetrac a "
'         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
'         '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'         .Source = .Source & "AND Left(a.coddetrac, 2)='" & dcCodEli & "'"
'         .Open
'      End With
'      If uorstcodetrac.RecordCount > 1 Then
'         MsgBox Choose(gsIdioma, "Existen Diarios Relacionados, No podra Eliminar", "Exist Relationed Journals, you could not eliminate them"), vbExclamation
'         With uorstcodetrac
'            .Close
'            .Source = "SELECT a.coddetrac, a.detdetrac, a.detdetracx "
'            .Source = .Source & "FROM codetrac a "
'            .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
'            '.Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'            .Open
'         End With
'         dgrMain.SetFocus
'         Exit Sub
'      End If
'   End If
'   ']
'fin 2015-06-26 proceso detrac
   'Mensaje de verificaci�n            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      uocnnMain.BeginTrans
      uorstMain.Delete
      uocnnMain.CommitTrans
   End If
   dgrMain.SetFocus
   With uorstcodetrac
      .Close
      .Source = "SELECT a.coddetrac, a.detdetrac, a.detdetracx "
      .Source = .Source & "FROM codetrac a "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      '.Source = .Source & "AND pdoano='" & gsAnoAct & "'"
      .Open
   End With
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
 '[Datos del formulario de impresi�n.  'Cambiar.
   frmLDetrac.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLDetrac.Show vbModal
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
   
'[ARREGLAR: B�squeda con distintos tipos de columna.
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
            .Item(dnNum).Caption = Choose(gsIdioma, "C�digo", "Code")
'            .Item(dnNum).Width = 600
             .Item(dnNum).Width = 100 * (uorstMain.Fields("coddetrac").DefinedSize + 4)
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Descripci�n", "Description")
'            .Item(dnNum).Width = 1500
             .Item(dnNum).Width = 100 * (uorstMain.Fields("detdetrac").DefinedSize)
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "Tasa", "Rate")
'            .Item(dnNum).Width = 1500
             .Item(dnNum).Width = 30 * (uorstMain.Fields("pctdetrac").DefinedSize)
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[C�digo propio del formulario.

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
