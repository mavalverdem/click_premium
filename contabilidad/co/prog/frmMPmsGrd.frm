VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMPmsGrd 
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
         Picture         =   "frmMPmsGrd.frx":0000
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
         Picture         =   "frmMPmsGrd.frx":014A
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
         Picture         =   "frmMPmsGrd.frx":0294
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
         Picture         =   "frmMPmsGrd.frx":0396
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
         Picture         =   "frmMPmsGrd.frx":0498
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
         Picture         =   "frmMPmsGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMPmsGrd"
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
Public uorstSGUsr As ADODB.Recordset
Public uorstSGMdl As ADODB.Recordset
Public uorstTGEmp As ADODB.Recordset
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
'   psConnStrgSele = "SELECT SGPms.CodUsr, b.RazEmp, c.DetMdl," _
                  & "  SGPms.CodEmp, SGPms.CodSis, SGPms.CodMdl," _
                  & "  SGPms.IndPms01, SGPms.IndPms02, SGPms.IndPms03, SGPms.IndPms04, SGPms.IndPms05, SGPms.IndPms06, SGPms.IndPms07, SGPms.IndPms08, SGPms.IndPms09, SGPms.IndPms10," _
                  & "  Concat(SGPms.CodUsr, SGPms.CodEmp, SGPms.CodMdl) as cLlave," _
                  & "  b.CodEmp, c.CodMdl," _
                  & "  SGPms.UsrCre, SGPms.FyHCre, SGPms.UsrMdf, SGPms.FyHMdf " _
                  & "FROM (SGPms" _
                  & "  LEFT JOIN TGEmp b ON SGPms.CodEmp=b.CodEmp)" _
                  & "  LEFT JOIN SGMdl c ON SGPms.CodMdl=c.CodMdl "
   psConnStrgSele = "SELECT SGPms.CodUsr, b.RazEmp, c.DetMdl," _
                  & "  SGPms.CodEmp, SGPms.CodSis, SGPms.CodMdl," _
                  & "  SGPms.IndPms01, SGPms.IndPms02, SGPms.IndPms03, SGPms.IndPms04, SGPms.IndPms05, SGPms.IndPms06, SGPms.IndPms07, SGPms.IndPms08, SGPms.IndPms09, SGPms.IndPms10," _
                  & "  Concat(SGPms.CodUsr, SGPms.CodEmp, SGPms.CodMdl) as cLlave," _
                  & "  SGPms.UsrCre, SGPms.FyHCre, SGPms.UsrMdf, SGPms.FyHMdf " _
                  & "FROM (SGPms" _
                  & "  LEFT JOIN TGEmp b ON SGPms.CodEmp=b.CodEmp)" _
                  & "  LEFT JOIN SGMdl c ON SGPms.CodMdl=c.CodMdl "
   psConnStrgOrde = "ORDER BY 1, 2, 3"

   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   Set uorstSGUsr = New ADODB.Recordset
   Set uorstSGMdl = New ADODB.Recordset
   Set uorstTGEmp = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "SGPms"
   End With
   With uorstSGUsr
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodUsr, NomUsr " _
              & "FROM SGUsr"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstSGMdl
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodMdl, DetMdl " _
              & "FROM SGMdl"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstTGEmp
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodEmp, RazEmp " _
              & "FROM TGEmp"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
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
   gpTUg_Nuevo Me, frmMPms             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
   
   If uorstMain.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If

   With frmMPms                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitaci�n de Llaves.       'Cambiar.
      .txtLlave(0).Enabled = False
      .txtLlave(1).Enabled = False
      .txtLlave(2).Enabled = False
      .cmdLlaveAyud(0).Enabled = False
      .cmdLlaveAyud(1).Enabled = False
      .cmdLlaveAyud(2).Enabled = False
      .lblLlaveDeta(0).Enabled = False
      .lblLlaveDeta(1).Enabled = False
      .lblLlaveDeta(2).Enabled = False
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
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If
   
   'Mensaje de verificaci�n            'Cambiar.
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
 '[Datos del formulario de impresi�n.  'Cambiar.
   frmLPms.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLPms.Show vbModal
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
      .Properties("Unique Table").Value = "SGPms"
   End With
   Set dgrMain.DataSource = uorstMain
   ppDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
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
            .Item(dnNum).Caption = "Usuario"
            .Item(dnNum).Width = 1500
         Case 1
            .Item(dnNum).Caption = "Empresa"
            .Item(dnNum).Width = 2950
         Case 2
            .Item(dnNum).Caption = "M�dulo"
            .Item(dnNum).Width = 3470
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
   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

