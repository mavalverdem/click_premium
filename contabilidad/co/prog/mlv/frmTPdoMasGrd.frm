VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmTPdoMasGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   4665
   ClientLeft      =   3945
   ClientTop       =   3630
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   7095
   Visible         =   0   'False
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7095
      TabIndex        =   8
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
         Picture         =   "frmTPdoMasGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
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
            TabIndex        =   6
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
         Left            =   6390
         Picture         =   "frmTPdoMasGrd.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmTPdoMasGrd.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmTPdoMasGrd.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmTPdoMasGrd.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTPdoMasGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   4035
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7117
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
End
Attribute VB_Name = "frmTPdoMasGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pnColumnaOrd As Integer
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2


Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
  On Error GoTo Err
  
  '[ARREGLAR. No acepta ordenar por columna de tablas secundarias en el recordset.
  If ColIndex = 1 Then Exit Sub
  ']ARREGLAR.
  
  pnColumnaOrd = ColIndex
  fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
  txtBuscar = ""
  
  frmTPdoGrd.usConnStrgOrde_CoPdoCprCta = "ORDER BY " & pnColumnaOrd + 1
  With frmTPdoGrd.uorstCoDPeCta
    .Close
    .Properties("Unique Table").Value = ps_Prefijo & "tmpcopdocprcta"
    .Source = frmTPdoGrd.usConnStrgSele_CoPdoCprCta & frmTPdoGrd.usConnStrgWher_CoPdoCprCta & frmTPdoGrd.usConnStrgOrde_CoPdoCprCta
    .Open
  End With
  Set dgrMain.DataSource = frmTPdoGrd.uorstCoDPeCta
  ppDatosGrid
  
  Exit Sub
Err:
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
  If frmTPdoGrd.uorstCoDPeCta.RecordCount = 0 Then Exit Sub
  Select Case KeyCode
   Case vbKeyHome
    frmTPdoGrd.uorstCoDPeCta.MoveFirst
   Case vbKeyEnd
    frmTPdoGrd.uorstCoDPeCta.MoveLast
  End Select
End Sub

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
      
  With frmTPdoGrd.uorstCoDPeCta
    If .State = adStateOpen Then .Close
    .Source = frmTPdoGrd.usConnStrgSele_CoPdoCprCta & frmTPdoGrd.usConnStrgWher_CoPdoCprCta & frmTPdoGrd.usConnStrgOrde_CoPdoCprCta
    .Open
    .Properties("Unique Table").Value = ps_Prefijo & "tmpcopdocprcta"
  End With
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  Me.Caption = Choose(gsIdioma, "Cuentas de Pedido", "Account Order")
  CaptionBotones Me, False, False, True, True, True, True, False, False, False, False, False, False, True, aLabel
  ']
  
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = frmTPdoGrd.uorstCoDPeCta
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
Public Sub cmdNuevo_Click()
  gpTUg_Nuevo Me, frmTPdoMas             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
  On Error GoTo Err
  
  If frmTPdoGrd.uorstCoDPeCta.RecordCount = 0 Then
  MsgBox TEXT_8001, vbCritical
  Exit Sub
  End If
  
  With frmTPdoMas                        'Cambiar Formulario de Datos.
  .zbNuevo = False
  .upDatosDesconectados 1
  '[Deshabilitación de Llaves.       'Cambiar.
  .txtDato(0).Enabled = False
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
  
  'Verificación de existencia de ítemes.
  If frmTPdoGrd.uorstCoDPeCta.RecordCount = 0 Then
    MsgBox TEXT_8001, vbCritical
    Exit Sub
  End If
   
  'Mensaje de verificación            'Cambiar.
  If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    frmTPdoGrd.uocnnMain.BeginTrans       'INICIA TRANSACCION.
    frmTPdoGrd.uorstCoDPeCta.Properties("Unique Table").Value = ps_Prefijo & "tmpcopdocprcta"
    frmTPdoGrd.uorstCoDPeCta.Delete
    frmTPdoGrd.uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.
  End If
  dgrMain.SetFocus
  Exit Sub
Err:
  gpErrores
  frmTPdoGrd.uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
'   gpTUg_Refrescar Me
   frmTPdoGrd.uorstCoDPeCta.Requery
   ppDatosGrid
   dgrMain.SetFocus
End Sub

Private Sub cmdSalir_Click()
  Dim sSentencia As String
  Dim nImporteMN As Double, nImporteME As Double, nImporteDF As Double
   
  ' Inicializo los datos de cuenta
  frmTPdo.txtDato(7).Text = ""
  frmTPdo.lblDatoDeta(7).Caption = ""
  frmTPdo.txtDato(8).Text = ""
  frmTPdo.lblDatoDeta(8).Caption = ""
  frmTPdo.txtDato(7).Enabled = True
  frmTPdo.cmdDatoAyud(7).Enabled = True
  frmTPdo.txtDato(8).Enabled = False
  frmTPdo.cmdDatoAyud(8).Enabled = False
  frmTPdo.cmdMas.Tag = INDMASCTA_INI
  If Not (frmTPdoGrd.uorstCoDPeCta.EOF And frmTPdoGrd.uorstCoDPeCta.BOF) And frmTPdoGrd.uorstCoDPeCta.RecordCount > 0 Then
    With frmTPdoGrd
      sSentencia = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impcta_mn), 0), 2) AS ImporteMN, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impcta_me), 0), 2) AS ImporteME, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impctadif), 0), 2) AS ImporteDF "
      sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcopdocprcta "
      Set .porstCancel = .uocnnMain.Execute(sSentencia)
      nImporteMN = CDec(.porstCancel!ImporteMN)
      nImporteME = CDec(.porstCancel!ImporteME)
      nImporteDF = CDec(.porstCancel!ImporteDF)
      .porstCancel.Close
    End With
    
    If (CDec(frmTPdo.txtDato(4).Text) <> CDec(nImporteMN)) Or (CDec(frmTPdo.txtDato(5).Text) <> CDec(nImporteME) Or (CDec(frmTPdo.txtDato(6).Text) <> CDec(nImporteDF))) Then
            MsgBox Choose(gsIdioma, "El importe total de las Cuentas Moneda Nacional : ", "The total amount of Acccounts National Currency :") & Format(nImporteMN, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser ", " and must be ") & frmTPdo.txtDato(4).Text & Chr(13) & Space(48) & Choose(gsIdioma, "Moneda Extranjera : ", "Foreign : ") & Format(nImporteME, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser ", " and must be ") & frmTPdo.txtDato(5).Text & Chr(13) & Space(48) & Choose(gsIdioma, "Diferencial : ", "Differential : ") & Format(nImporteDF, FORMATO_NUM_1) & Choose(gsIdioma, " y debe ser ", " and must be ") & frmTPdo.txtDato(6).Text, vbCritical
            If MsgBox("Desea Actualizar ", vbInformation + vbYesNo + vbDefaultButton1, "Sistema de Contabilidad") = vbYes Then
                frmTPdo.txtDato(4).Text = nImporteMN
                frmTPdo.txtDato(5).Text = nImporteME
                frmTPdo.txtDato(6).Text = nImporteDF
            End If
            Exit Sub
    End If
    
    frmTPdoGrd.uorstCoDPeCta.MoveFirst
    ' Pintado de primer registro
    frmTPdo.txtDato(7).Text = frmTPdoGrd.uorstCoDPeCta!codcta
    frmTPdo.txtDato(8).Text = IIf(IsNull(frmTPdoGrd.uorstCoDPeCta!codcco), "", frmTPdoGrd.uorstCoDPeCta!codcco)
    frmTPdo.txtDato(7).Enabled = False
    frmTPdo.txtDato(8).Enabled = False
    frmTPdo.cmdDatoAyud(7).Enabled = False
    frmTPdo.cmdDatoAyud(8).Enabled = False
    frmTPdo.cmdMas.Tag = INDMASCTA_MAS
    frmTPdo.cmdMas.Enabled = True
  End If
  Unload Me
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With frmTPdoGrd.uorstCoDPeCta
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
      frmTPdoGrd.uorstCoDPeCta.Bookmark = dvRegistroActual
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
        .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
        .Item(dnNum).Width = 100 * (frmTPdoGrd.uorstCoDPeCta.Fields("codcta").DefinedSize + 1)
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "C.Costo", "C.Center")
        .Item(dnNum).Width = 100 * (frmTPdoGrd.uorstCoDPeCta.Fields("codcco").DefinedSize + 1)
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "Importe ", "Amount ") & TPOMON_NAC_TXT_0
        .Item(dnNum).Width = 1200
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
        .Item(dnNum).Alignment = dbgRight
       Case 3
        .Item(dnNum).Caption = Choose(gsIdioma, "Importe ", "Amount ") & TPOMON_EXT_TXT_0
        .Item(dnNum).Width = 1200
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
        .Item(dnNum).Alignment = dbgRight
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
   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

