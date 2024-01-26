VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmTConGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   1860
   ClientTop       =   2010
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8775
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
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
      ScaleWidth      =   8775
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
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
         Picture         =   "frmTConGrd.frx":0000
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
         Picture         =   "frmTConGrd.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
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
         Picture         =   "frmTConGrd.frx":0204
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
         CausesValidation=   0   'False
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
         Left            =   8055
         Picture         =   "frmTConGrd.frx":034E
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
         Picture         =   "frmTConGrd.frx":0498
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
         Picture         =   "frmTConGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTConGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uocnnNoGrabable As ADODB.Connection
Public uorstMain As ADODB.Recordset
Public uorstMain_Grd As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgSele_Grd As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer

'[Propio del formulario.
Public uorstTGAux As ADODB.Recordset
Public uorstTGTCb As ADODB.Recordset
Public porstCancel As ADODB.Recordset
']

Private Sub cmdImprimir_Click(Index As Integer)
  If uorstMain.RecordCount = 0 Then
     MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
     Exit Sub
  End If
 '[Datos del formulario de impresión.  'Cambiar.
   frmLCon.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLCon.Show vbModal
 ']
End Sub

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
  psConnStrgSele_Grd = "SELECT coconser.codcon, coconser.codaux, b.razaux, coconser.Fehcon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "coconser." & Choose(gsIdioma, "detcon", "detconx") & " AS cdetcon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "coconser.tpomon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE coconser.tpomon WHEN '" & TPOMON_NAC & "' THEN coconser.impmn ELSE coconser.impme END) AS cImporte, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(coconser.codcon, coconser.codaux)", "(coconser.codcon+coconser.codaux)") & " AS cLlave "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM coconser "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGAux b ON coconser.codemp=b.codemp AND coconser.codaux=b.codaux "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "WHERE coconser.codemp='" & gsCodEmp & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND coconser.pdoano='" & gsAnoAct & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND coconser.mespvs='" & gsMesAct & "' "
  
  psConnStrgSele = "SELECT coconser.codcon, coconser.codaux, "
  psConnStrgSele = psConnStrgSele & "coconser.fehcon, coconser.detcon, coconser.detconx, "
  psConnStrgSele = psConnStrgSele & "coconser.tpomon, coconser.imptcb, coconser.impmn, coconser.impme, "
  psConnStrgSele = psConnStrgSele & "coconser.UsrCre, coconser.FyHCre, coconser.UsrMdf, coconser.FyHMdf, "
  psConnStrgSele = psConnStrgSele & "coconser.codemp, coconser.pdoano, coconser.mespvs, "
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "Concat(coconser.codcon, coconser.codaux)", "(coconser.codcon+coconser.codaux)") & " AS cLlave "
  psConnStrgSele = psConnStrgSele & "FROM coconser "
  psConnStrgSele = psConnStrgSele & "WHERE coconser.codemp='" & gsCodEmp & "' "
  psConnStrgSele = psConnStrgSele & "AND coconser.pdoano='" & gsAnoAct & "' "
  psConnStrgSele = psConnStrgSele & "AND coconser.mespvs='" & gsMesAct & "' "
  psConnStrgOrde = "ORDER BY coconser.codcon, coconser.codaux"
  
  Set uocnnMain = New ADODB.Connection
  Set uocnnNoGrabable = New ADODB.Connection
  Set uorstMain = New ADODB.Recordset
  Set uorstMain_Grd = New ADODB.Recordset
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTCb = New ADODB.Recordset
  Set porstCancel = New ADODB.Recordset
  With uocnnMain
     .CursorLocation = adUseClient
     .ConnectionString = CONNSTRG & gsNomBDS
     .Open
  End With
  With uocnnNoGrabable
     .CursorLocation = adUseClient
     .ConnectionString = CONNSTRG & gsNomBDS
     .Open
  End With
  With uorstMain_Grd
     .ActiveConnection = uocnnMain
     .Source = psConnStrgSele_Grd & psConnStrgOrde
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "coconser"
  End With
  With uorstMain
     .ActiveConnection = uocnnMain
     .Source = psConnStrgSele & psConnStrgOrde
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "coconser"
  End With
  With uorstTGTCb
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.FehTCb, a.ImpTCb_Cpr, a.ImpTCb_Vta "
     .Source = .Source & "FROM TGTCb a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
  With porstCancel
     .ActiveConnection = uocnnMain
     .CursorType = adOpenDynamic
     .LockType = adLockBatchOptimistic ' adLockOptimistic
  End With
  With uorstTGAux
     .ActiveConnection = uocnnNoGrabable
     .Source = "SELECT a.CodAux, a.RazAux "
     .Source = .Source & "FROM TGAux a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.IndCli=1 "
     .Source = .Source & "AND a.EstAux='" & ESTAUX_ACT & "'"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  ']
  
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = uorstMain_Grd
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  ']
End Sub

Private Sub Form_Activate()
  'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
  zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
  upDatosGrid
  fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub
Private Sub Form_Resize()
   On Error Resume Next
  
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   uorstTGAux.Close
   uorstTGTCb.Close
   uorstMain_Grd.Close
   uorstMain.Close
   uocnnMain.Close
   Set porstCancel = Nothing
   Set uorstTGAux = Nothing
   Set uorstTGTCb = Nothing
   Set uorstMain_Grd = Nothing
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Private Sub cmdNuevo_Click()
  '[Propio del formulario.
  'Verificación de Mes Cerrado.
  If gbCieCpr Then MsgBox TEXT_9016, vbCritical: Exit Sub
  ']
  gpTUg_Nuevo Me, frmTCon             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
  On Error GoTo Err
  
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub

  '[Búsqueda del ítem.
  uorstMain.Requery
  uorstMain.MoveFirst
  uorstMain.Find "cLlave='" & uorstMain_Grd!codcon & uorstMain_Grd!codaux & "'"
  ']

  With frmTCon                        'Cambiar Formulario de Datos.
    .zbNuevo = False
    .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
    .txtLlave(0).Enabled = False
    .txtDato(0).Enabled = False
    .cmdDatoAyud(0).Enabled = False
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
  Dim dsLlaveSiguiente As String
  
  On Error GoTo Err
  
  'Verificación de Mes Cerrado.
  If gbCieCpr Then MsgBox TEXT_9016, vbCritical: Exit Sub
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub
   
  ' Mensaje de verificación            'Cambiar.
  If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & "-" & Trim(dgrMain.Columns(2)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    With porstCancel
      .Source = "SELECT mespvs, codcon, codaux "
      .Source = .Source & "FROM cocprdoc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND mespvs='" & gsMesAct & "' AND codaux='" & uorstMain_Grd!codaux & "' "
      .Source = .Source & "AND codcon='" & uorstMain_Grd!codcon & "' "
      .Source = .Source & "UNION "
      .Source = .Source & "SELECT mespvs, codcon, codaux "
      .Source = .Source & "FROM cohprdoc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND mespvs='" & gsMesAct & "' AND codaux='" & uorstMain_Grd!codaux & "' "
      .Source = .Source & "AND codcon='" & uorstMain_Grd!codcon & "' "
      .Source = .Source & "UNION "
      .Source = .Source & "SELECT mespvs, codcon, codaux "
      .Source = .Source & "FROM covtadoc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND mespvs='" & gsMesAct & "' AND codaux='" & uorstMain_Grd!codaux & "' "
      .Source = .Source & "AND codcon='" & uorstMain_Grd!codcon & "'"
      .Open
      If porstCancel.RecordCount = 0 Then
        uorstMain.MoveFirst
        uorstMain.Find "cLlave = '" & uorstMain_Grd!codcon & uorstMain_Grd!codaux & "'"
        
        uocnnMain.BeginTrans       'INICIA TRANSACCION.
        uorstMain.Properties("Unique Table").Value = "coconser"
        uorstMain.Delete
        uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

        'Busca siguiente ítem.
        With uorstMain_Grd
          .MoveNext
          If .EOF Then .MoveLast
          dsLlaveSiguiente = !codcon & !codaux
          .Requery
          If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
        End With
        upDatosGrid
        ' actualizo recordset principal
        uorstMain.Requery
        If uorstMain.RecordCount > 0 Then uorstMain.Find "cLlave = '" & dsLlaveSiguiente & "'"
      Else
        MsgBox Choose(gsIdioma, "Debe eliminar antes las Provisiones.", " The Provisions must be eliminated before."), vbExclamation
      End If
    End With
    porstCancel.Close
  End If
  dgrMain.SetFocus
  Exit Sub
Err:
  gpErrores
  
  uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
'[ARREGLAR. Usar gpTUg_Refrescar Me, pero se debe cambiar ppDatosGrid a upDatosGrid para todos los _
            formularios que lo usan (formularios de registro único).
''   gpTUg_Refrescar Me
   uorstMain_Grd.Requery
   upDatosGrid
   
   dgrMain.SetFocus
']ARREGLAR.
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
'[ARREGLAR. No acepta ordenar por columna de tablas secundarias en el recordset.
   If ColIndex = 3 Then Exit Sub
']ARREGLAR.

   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = "ORDER BY "
   Select Case pnColumnaOrd            'Cambiar.
'   Case 1
'      psConnStrgOrde = psConnStrgOrde & "2, 3, 4"
   Case Else
      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
   End Select
   With uorstMain_Grd
      .Close
      .Properties("Unique Table").Value = "coconser"
      .Source = psConnStrgSele_Grd & psConnStrgOrde
      .Open
   End With
   Set dgrMain.DataSource = uorstMain_Grd
   upDatosGrid

   Exit Sub
Err:
   gpErrores
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

Public Sub upDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Servicio", "Service")
            .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("codcon").DefinedSize + 1)
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("codaux").DefinedSize + 0.5)
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
            .Item(dnNum).Width = 1750
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "F.Emisión", "Issue Date")
            .Item(dnNum).Width = 980
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "Detalle", "Detail")
            .Item(dnNum).Width = 1750
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Mon", "Cur")
            .Item(dnNum).Width = 250
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe", "Amount")
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
   cmdImprimir(1).Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property
