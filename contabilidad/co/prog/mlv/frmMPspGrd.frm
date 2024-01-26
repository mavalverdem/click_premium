VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMPspGrd 
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
   Begin MSComctlLib.ImageList Imagenes 
      Left            =   6240
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMPspGrd.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
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
      Begin MSComctlLib.Toolbar ToolbarCostos 
         Height          =   600
         Left            =   3615
         TabIndex        =   10
         Top             =   -15
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1058
         ButtonWidth     =   1058
         ButtonHeight    =   953
         ImageList       =   "Imagenes"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Costos"
               ImageIndex      =   1
               Style           =   5
            EndProperty
         EndProperty
      End
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
         Picture         =   "frmMPspGrd.frx":015A
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
         Left            =   4440
         TabIndex        =   0
         Top             =   0
         Width           =   1935
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
         Picture         =   "frmMPspGrd.frx":02A4
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
         Picture         =   "frmMPspGrd.frx":03EE
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
         Picture         =   "frmMPspGrd.frx":04F0
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
         Picture         =   "frmMPspGrd.frx":05F2
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
         Picture         =   "frmMPspGrd.frx":06F4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMPspGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain As ADODB.Recordset
Public porstCOCta As ADODB.Recordset
Public porstUltOrdRep As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer
Public porstCOCCo As ADODB.Recordset '2016-07-12 correccion de presupuesto

'[Propio del formulario.
'Public uorstTGDtt As ADODB.Recordset
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
 
   psConnStrgSele = "SELECT COPsp.OrdRep, COPsp.CodCta, COPsp.CodCco," & Choose(gsIdioma, "COCta.DetCta", "COCta.DetCtax") & " AS DetCta, "
   psConnStrgSele = psConnStrgSele & "COPsp.ImpMN_01, COPsp.ImpMN_02, COPsp.ImpMN_03, COPsp.ImpMN_04, COPsp.ImpMN_05, COPsp.ImpMN_06, COPsp.ImpMN_07, COPsp.ImpMN_08, COPsp.ImpMN_09, COPsp.ImpMN_10, COPsp.ImpMN_11, COPsp.ImpMN_12, "
   psConnStrgSele = psConnStrgSele & "COPsp.ImpME_01, COPsp.ImpME_02, COPsp.ImpME_03, COPsp.ImpME_04, COPsp.ImpME_05, COPsp.ImpME_06, COPsp.ImpME_07, COPsp.ImpME_08, COPsp.ImpME_09, COPsp.ImpME_10, COPsp.ImpME_11, COPsp.ImpME_12, "
   psConnStrgSele = psConnStrgSele & "COPsp.codemp, COPsp.pdoano, "
   psConnStrgSele = psConnStrgSele & "COPsp.UsrCre, COPsp.FyHCre, COPsp.UsrMdf, COPsp.FyHMdf, Concat(COPsp.CodCta, COPsp.CodCco) AS llave "
   psConnStrgSele = psConnStrgSele & "FROM COPsp "
   psConnStrgSele = psConnStrgSele & "LEFT JOIN COCta As COCta ON COPsp.codemp=COCta.codemp AND COPsp.pdoano=COCta.pdoano AND COPsp.CodCta=COCta.CodCta "
   psConnStrgSele = psConnStrgSele & "WHERE COPsp.codemp='" & gsCodEmp & "' "
   psConnStrgSele = psConnStrgSele & "AND COPsp.pdoano='" & gsAnoAct & "' "
   psConnStrgOrde = "ORDER BY COPsp.OrdRep, COPsp.CodCta"

   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstCOCCo = New ADODB.Recordset '2016-07-12 correccion de presupuesto
   Set porstUltOrdRep = New ADODB.Recordset
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
      .Properties("Unique Table").Value = "COPsp"
   End With
   With porstCOCta
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta, TpoCta "
      .Source = .Source & "FROM COCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
'      .Properties("Unique Table").Value = "COCta"
   End With
'ini 2016-07-12 correccion de presupuesto
   With porstCOCCo
      .ActiveConnection = uocnnMain
      .Source = "SELECT codcco, " & Choose(gsIdioma, "detcco", "detccox") & " AS DetCCo "
      .Source = .Source & "FROM cocco "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
'      .Properties("Unique Table").Value = "COCta"
   End With
'fin 2016-07-12 correccion de presupuesto
   
   With porstUltOrdRep
      .ActiveConnection = uocnnMain
      .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(OrdRep), 0) AS cUltOrd "
      .Source = .Source & "FROM COPsp "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
   End With
   
 ']
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  ']
   
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain
   
'*****************************
Dim Rst As New Recordset
Dim sql As String
Dim i As Integer
sql = "select codcco,detcco from cocco where codemp='" & gsCodEmp & "' and pdoano='" & gsAnoAct & "' and LENGTH(codcco)=2"
Rst.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
Rst.MoveFirst
i = 1
While Not Rst.EOF
    ToolbarCostos.Buttons(1).ButtonMenus.Add(i).Key = "A" & i
    ToolbarCostos.Buttons(1).ButtonMenus(i).Text = Rst(0) & " " & Rst(1)
    i = i + 1
    Rst.MoveNext
Wend
ToolbarCostos.Buttons(1).ButtonMenus.Add(i).Key = "A" & i
ToolbarCostos.Buttons(1).ButtonMenus(i).Text = "Todos"
'*****************************
   
   
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
   porstCOCta.Close
   porstCOCCo.Close '2016-07-12 correccion de presupuesto
   uorstMain.Close
   uocnnMain.Close
   Set uorstMain = Nothing
   Set porstUltOrdRep = Nothing
   Set porstCOCta = Nothing
   Set porstCOCCo = Nothing '2016-07-12 correccion de presupuesto
   Set uocnnMain = Nothing
End Sub

Public Sub cmdNuevo_Click()
   gpTUg_Nuevo Me, frmMPsp             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
   
   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   With frmMPsp                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
      .txtLlave(0).Enabled = False
'ini 2016-07-12 correccion de presupuesto
      .txtLlave(1).Enabled = False
'fin 2016-07-12 correccion de presupuesto
      
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
'ini 2016-07-12 correccion de presupuesto
'   On Error GoTo Err
'fin 2016-07-12 correccion de presupuesto

   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?" & " (" & Trim(dgrMain.Columns(2)) & ")?  ", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      uocnnMain.BeginTrans
'ini 2016-07-12 correccion de presupuesto
'      uorstMain.Delete
    uocnnMain.Execute "DELETE FROM COPsp " _
    & "WHERE CodEmp='" & gsCodEmp & "'  AND PdoAno='" & gsAnoAct & "' " _
    & " AND CodCta='" & Trim(dgrMain.Columns(1)) & "'  AND CodCco='" & Trim(dgrMain.Columns(2)) & "' " _
    & " AND OrdRep='" & Trim(dgrMain.Columns(0)) & "' "
    uorstMain.Requery
'fin 2016-07-12 correccion de presupuesto
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
   frmLPsp.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLPsp.Show vbModal
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
Private Sub ToolbarCostos_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
Case "A" & ButtonMenu.Index
    
   psConnStrgSele = "SELECT COPsp.OrdRep, COPsp.CodCta, COPsp.CodCco," & Choose(gsIdioma, "COCta.DetCta", "COCta.DetCtax") & " AS DetCta, "
   psConnStrgSele = psConnStrgSele & "COPsp.ImpMN_01, COPsp.ImpMN_02, COPsp.ImpMN_03, COPsp.ImpMN_04, COPsp.ImpMN_05, COPsp.ImpMN_06, COPsp.ImpMN_07, COPsp.ImpMN_08, COPsp.ImpMN_09, COPsp.ImpMN_10, COPsp.ImpMN_11, COPsp.ImpMN_12, "
   psConnStrgSele = psConnStrgSele & "COPsp.ImpME_01, COPsp.ImpME_02, COPsp.ImpME_03, COPsp.ImpME_04, COPsp.ImpME_05, COPsp.ImpME_06, COPsp.ImpME_07, COPsp.ImpME_08, COPsp.ImpME_09, COPsp.ImpME_10, COPsp.ImpME_11, COPsp.ImpME_12, "
   psConnStrgSele = psConnStrgSele & "COPsp.codemp, COPsp.pdoano, "
   psConnStrgSele = psConnStrgSele & "COPsp.UsrCre, COPsp.FyHCre, COPsp.UsrMdf, COPsp.FyHMdf, Concat(COPsp.CodCta, COPsp.CodCco) AS llave "
   psConnStrgSele = psConnStrgSele & "FROM COPsp "
   psConnStrgSele = psConnStrgSele & "LEFT JOIN COCta As COCta ON COPsp.codemp=COCta.codemp AND COPsp.pdoano=COCta.pdoano AND COPsp.CodCta=COCta.CodCta "
   psConnStrgSele = psConnStrgSele & "WHERE COPsp.codemp='" & gsCodEmp & "' "
   psConnStrgSele = psConnStrgSele & "AND COPsp.pdoano='" & gsAnoAct & "' "
   If ButtonMenu.Text <> "Todos" Then
    psConnStrgSele = psConnStrgSele & "AND COPsp.codcco='" & Left(ButtonMenu.Text, 2) & "'"
   End If
   uorstMain.Close

   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "COPsp"
   End With
    
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain
   ppDatosGrid
    
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
            .Item(dnNum).Caption = Choose(gsIdioma, "Orden", "Order")
            .Item(dnNum).Width = 600
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
            .Item(dnNum).Width = 800
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Costos", "")
            .Item(dnNum).Width = 800
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "Descripción", "Description")
            .Item(dnNum).Width = 4600
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_01_MN", "Amount_01_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_02_MN", "Amount_02_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_03_MN", "Amount_03_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_04_MN", "Amount_04_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 8
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_05_MN", "Amount_05_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 9
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_06_MN", "Amount_06_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 10
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_07_MN", "Amount_07_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 11
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_08_MN", "Amount_08_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 12
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_09_MN", "Amount_09_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 13
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_10_MN", "Amount_10_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 14
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_11_MN", "Amount_11_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
         Case 15
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe_12_MN", "Amount_12_NC")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1
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

