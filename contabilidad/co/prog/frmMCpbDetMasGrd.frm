VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMCpbDetMasGrd 
   Caption         =   "[Entidad Tipo Asiento]"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleMode       =   0  'User
   ScaleWidth      =   10756.2
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7095
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
         Picture         =   "frmMCpbDetMasGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmMCpbDetMasGrd.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmMCpbDetMasGrd.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmMCpbDetMasGrd.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   720
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
         Picture         =   "frmMCpbDetMasGrd.frx":0408
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
         TabIndex        =   2
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   200
            Width           =   2415
         End
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
         Picture         =   "frmMCpbDetMasGrd.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   1
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
      Caption         =   "Flujo de Caja"
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
            ColumnWidth     =   1523.856
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1523.856
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMCpbDetMasGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'sirve para darle el mes 00
'por defecto todo hira ahi
Public rcMesAct

Public uorstMain As ADODB.Recordset
Public uorstMain_Grd As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgSele_Grd As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer

'[Propio del formulario.
Public uorstCOFjo As ADODB.Recordset
']

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
'[ARREGLAR. No acepta ordenar por columna de tablas secundarias en el recordset.
   If ColIndex = 1 Then Exit Sub
']ARREGLAR.

   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = "ORDER BY " & pnColumnaOrd + 1
   With uorstMain_Grd
      .Close
      .Properties("Unique Table").Value = "comacpbdetFjo"
      .Source = psConnStrgSele_Grd & psConnStrgOrde
      .Open
   End With
   Set dgrMain.DataSource = uorstMain_Grd
   ppDatosGrid

   Exit Sub
Err:
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   If uorstMain_Grd.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      uorstMain_Grd.MoveFirst
   Case vbKeyEnd
      uorstMain_Grd.MoveLast
   End Select
End Sub

Private Sub Form_Load()
'rcMesAct = "00"
'2014-03-31 error de validacion fecha cance prov
'rcMesAct = "01"
rcMesAct = "12"

 '[Recordsets                          'Cambiar.
   psConnStrgSele_Grd = "SELECT " & ps_Prefijo & "tmpcomacpbdetFjo.CodFjo, " & Choose(gsIdioma, "CoFjo.DetFjo", "CoFjo.DetFjox") & ", ImpMN, ImpME, MesPvs, CodDro, NroCpb, "
   psConnStrgSele_Grd = psConnStrgSele_Grd & "NroIte, NroOrd, CodCta, TpoCtb, "
   psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "CONCAT(MesPvs, CodDro, NroCpb, NroIte, NroOrd, tmpcomacpbdetFjo.CodFjo)", "(MesPvs+CodDro+NroCpb+RTrim(NroIte)+RTrim(NroOrd)+#tmpcomacpbdetFjo.CodFjo)") & " AS cLlave, "
   psConnStrgSele_Grd = psConnStrgSele_Grd & ps_Prefijo & "tmpcomacpbdetFjo.UsrCre, " & ps_Prefijo & "tmpcomacpbdetFjo.FyHCre, " & ps_Prefijo & "tmpcomacpbdetFjo.UsrMdf, " & ps_Prefijo & "tmpcomacpbdetFjo.FyHMdf, "
   psConnStrgSele_Grd = psConnStrgSele_Grd & ps_Prefijo & "tmpcomacpbdetFjo.codemp, " & ps_Prefijo & "tmpcomacpbdetFjo.pdoano "
   psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM (" & ps_Prefijo & "tmpcomacpbdetFjo "
   psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN CoFjo ON " & ps_Prefijo & "tmpcomacpbdetFjo.codemp=CoFjo.codemp AND " & ps_Prefijo & "tmpcomacpbdetFjo.pdoano=CoFjo.pdoano AND " & ps_Prefijo & "tmpcomacpbdetFjo.CodFjo=CoFjo.CodFjo) "
   psConnStrgOrde = "ORDER BY 1"
   
   psConnStrgSele = "SELECT MesPvs, CodDro, NroCpb, NroIte, NroOrd, "
   psConnStrgSele = psConnStrgSele & "CodFjo, CodCta, TpoCtb, ImpMN, ImpME, "
   psConnStrgSele = psConnStrgSele & "UsrCre, FyHCre, UsrMdf, FyHMdf, codemp, pdoano "
   psConnStrgSele = psConnStrgSele & "FROM " & ps_Prefijo & "tmpcomacpbdetFjo "

   Set uorstMain_Grd = New ADODB.Recordset
   Set uorstMain = New ADODB.Recordset
   Set uorstCOFjo = New ADODB.Recordset
   With uorstMain
      .ActiveConnection = frmMCpbGrd.uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = ps_Prefijo & "tmpcomacpbdetFjo"
   End With
   With uorstMain_Grd
      .ActiveConnection = frmMCpbGrd.uocnnMain
      .Source = psConnStrgSele_Grd & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = ps_Prefijo & "tmpcomacpbdetFjo"
   End With
   With uorstCOFjo
      .ActiveConnection = frmMCpbGrd.uocnnMain
     .Source = "SELECT a.CodFjo, " & Choose(gsIdioma, "a.DetFjo", "a.DetFjox") & " AS DetFjo "
     .Source = .Source & "FROM COFjo a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
     .Source = .Source & "AND a.pdoano='" & gsAnoAct & "'"
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodFjo)>2"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenKeyset
      .LockType = adLockOptimistic
      .Open
   End With
 ']
      
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain_Grd
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
   uorstMain_Grd.Close
   Set uorstMain = Nothing
   Set uorstMain_Grd = Nothing
End Sub

Public Sub cmdNuevo_Click()
   gpTUg_Nuevo Me, frmMCpbDetMas             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
   
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

 '[Búsqueda del ítem.
   uorstMain.Requery
   uorstMain.MoveFirst
   uorstMain.Find "NroOrd='" & uorstMain_Grd!NroOrd & "'"
 ']
   With frmMCpbDetMas                        'Cambiar Formulario de Datos.
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
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If
   
   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
     uorstMain.MoveFirst
     uorstMain.Find "NroOrd = '" & uorstMain_Grd!NroOrd & "'"
    
     frmMCpbGrd.uocnnMain.BeginTrans       'INICIA TRANSACCION.
     uorstMain.Properties("Unique Table").Value = ps_Prefijo & "tmpcomacpbdetFjo"
     uorstMain.Delete
     frmMCpbGrd.uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.
    
     'Busca siguiente ítem.
     With uorstMain_Grd
       .MoveNext
       If .EOF Then .MoveLast
       .Requery
       If .RecordCount > 0 Then .Find "NroOrd = '" & uorstMain_Grd!NroOrd & "'"
     End With
     ppDatosGrid
   End If
   dgrMain.SetFocus
   Exit Sub
Err:
   gpErrores
   
   frmMCpbGrd.uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
'   gpTUg_Refrescar Me
   uorstMain_Grd.Requery
   ppDatosGrid
   dgrMain.SetFocus
End Sub

Private Sub cmdSalir_Click()
   
  Static nTotalMN, nTotalME As Currency
  Static uorstTmp As ADODB.Recordset
  
  nTotalMN = CDec(frmMCpbDet.txtImporte(0).Text) + CDec(frmMCpbDet.txtImporte(1).Text)
  nTotalME = CDec(frmMCpbDet.txtImporte(2).Text) + CDec(frmMCpbDet.txtImporte(3).Text)
  ' Inicializo los datos de flujo
  frmMCpbDet.txtDato(9).Text = ""
  frmMCpbDet.lblDatoDeta(4).Caption = ""
  frmMCpbDet.txtDato(9).Enabled = True
  frmMCpbDet.cmdDatoAyud(5).Enabled = True
  frmMCpbDet.cmdMasFjo.Tag = 0
  If Not (uorstMain_Grd.EOF And uorstMain_Grd.BOF) Then
    Set uorstTmp = New ADODB.Recordset
    With uorstTmp
      .ActiveConnection = frmMCpbGrd.uocnnMain
      .Source = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(ImpMN), 0), 2) AS nImpTotMN, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(ImpME), 0), 2) AS nImpTotME "
      .Source = .Source & "FROM " & ps_Prefijo & "tmpcomacpbdetFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND MesPvs='" & rcMesAct & "' "
      .Source = .Source & "AND CodDro='" & frmMCpbCab.txtLlave(0).Text & "' "
      .Source = .Source & "AND NroCpb='" & frmMCpbCab.txtLlave(1).Text & "' "
      .Source = .Source & "AND NroIte=" & frmMCpbDet.pnNroIte
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
      .Open
       If (CDec(nTotalMN) <> CDec(!nImpTotMN)) Or (CDec(nTotalME) <> CDec(!nImpTotME)) Then
         MsgBox "El importe total de los Flujos Moneda Nacional : " & Format(!nImpTotMN, FORMATO_NUM_1) & " y debe ser " & Format(nTotalMN, FORMATO_NUM_1) & Chr(13) & Space(48) & "Moneda Extranjera : " & Format(!nImpTotME, FORMATO_NUM_1) & " y debe ser " & Format(nTotalME, FORMATO_NUM_1), vbCritical
         .Close
         Exit Sub
       End If
      .Close
    End With
    uorstMain.MoveFirst
    ' Pintado de primer flujo de caja
    frmMCpbDet.txtDato(9).Text = uorstMain!CodFjo
    frmMCpbDet.txtDato(9).Enabled = False
    frmMCpbDet.cmdDatoAyud(5).Enabled = False
    frmMCpbDet.cmdMasFjo.Tag = 2
    frmMCpbDet.cmdMasFjo.Enabled = True
  End If
  Set uorstTmp = Nothing
   
  Unload Me
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With uorstMain_Grd
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
      uorstMain_Grd.Bookmark = dvRegistroActual
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
           .Item(dnNum).Caption = "Flujo"
           .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("CodFjo").DefinedSize + 4)
         Case 1
           .Item(dnNum).Caption = "Descripción"
           .Item(dnNum).Width = 68 * (uorstMain_Grd.Fields("DetFjo").DefinedSize)
         Case 2
           .Item(dnNum).Caption = "Importe " & TPOMON_NAC_TXT_0
           .Item(dnNum).Width = 1150
           .Item(dnNum).Alignment = dbgRight
           .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 3
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



