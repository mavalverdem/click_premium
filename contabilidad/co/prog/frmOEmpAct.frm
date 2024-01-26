VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOEmpAct 
   Caption         =   "[Entidad]"
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   1350
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame frmPeriodo 
      Caption         =   " Periodo "
      ForeColor       =   &H00000080&
      Height          =   2655
      Left            =   5220
      TabIndex        =   7
      Top             =   570
      Width           =   1560
      Begin VB.ListBox lstAnoAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   780
         ItemData        =   "frmOEmpAct.frx":0000
         Left            =   225
         List            =   "frmOEmpAct.frx":009A
         TabIndex        =   9
         Top             =   450
         Width           =   1050
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1500
         Width           =   1350
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   690
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   1260
         Width           =   630
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4683
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
      ScaleWidth      =   6825
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6825
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
         Left            =   720
         Picture         =   "frmOEmpAct.frx":01CA
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   1500
         TabIndex        =   0
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   4
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
         Left            =   5415
         Picture         =   "frmOEmpAct.frx":0314
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
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
         Picture         =   "frmOEmpAct.frx":045E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOEmpAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstMain As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer

Private psAnoAct As String

'[Propio del formulario.
Private porstSGPms As ADODB.Recordset
Private porstCoCfg As ADODB.Recordset
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   psConnStrgSele = "SELECT DISTINCT a.codemp, a.razEmp, a.rucemp, "
   psConnStrgSele = psConnStrgSele & "a.localidademp, a.direccion, a.actividademp, a.repapepaterno, a.repapematerno, "
   psConnStrgSele = psConnStrgSele & "a.repnombre, a.repdocumento, a.conapepaterno, a.conapematerno, a.connombre, a.condocumento "
   psConnStrgSele = psConnStrgSele & "FROM tgemp a, sgpms b "
   psConnStrgSele = psConnStrgSele & "WHERE b.codemp=a.codemp AND b.codusr='" & gsCodUsr & "' "
   psConnStrgOrde = "ORDER BY 1"
   
   Set pocnnMain = New ADODB.Connection
   Set porstMain = New ADODB.Recordset
   Set porstSGPms = New ADODB.Recordset
   Set porstCoCfg = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With porstMain
      .ActiveConnection = pocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
'      .Properties("Unique Table").Value = "TGEmp"
      .Find "CodEmp='" & gsCodEmp & "'"
   End With
   With porstSGPms
      .ActiveConnection = pocnnMain
      .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(CodUsr, CodEmp, CodSis)", "(CodUsr+CodEmp+CodSis)") & " AS cLlave "
      .Source = .Source & "FROM SGPms"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstCoCfg
    .Source = "SELECT pdoano, MesAtu, TpoMon_Fnc, IndMNE, "
    .Source = .Source & "CodCta_Nv3, CodCta_Nv4, CodCta_Nv5, CodCta_Nv6, CodCta_Nv7, CodCta_Nv8, "
    .Source = .Source & "CodTDc_Pcp, CodTDc_Rtc, CodCta_Pcp, CodCta_Rtc, IndRtc, IndPcp, "
    .Source = .Source & "CodCCo_Nv3, CodCCo_Nv5, TpoGlo_Rtc, GloDocr_Rtc, GloDocn_Rtc, "
    .Source = .Source & "coddro_ing, coddro_egr, ejerfran, prodestino "
    .Source = .Source & "FROM CoCfg "
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
   End With
 ']
   
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = porstMain

   Dim n_Index As Integer
   For n_Index = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      cmbPeriodo.AddItem Choose(n_Index + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
    Else
      cmbPeriodo.AddItem Choose(n_Index + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
    End If
   Next n_Index
   cmbPeriodo.ListIndex = Val(gsMesAct)

  '[ Cargo mensajes de botones y etiquetas
  ReDim aLabel(2, 2)
  frmPeriodo.Caption = Choose(gsIdioma, "Periodo Contable", "Accounting Period")
  For n_Index = 0 To 1
    aLabel(n_Index, 0) = Choose(n_Index + 1, "Ejercicio", "Periodo")
    aLabel(n_Index, 1) = Choose(n_Index + 1, "Fiscal year", "Period")
  Next n_Index
  CaptionBotones Me, True, False, False, False, False, True, False, False, False, False, False, False, True, aLabel
  ']

   psAnoAct = gsAnoAct
   lstAnoAct.ListIndex = (lstAnoAct.ListCount - 1) - (Val(psAnoAct) - 2001)
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
  
'   gpTUg_Resize Me
  'Esto cambiará el tamaño de la cuadrícula al cambiar el tamaño del formulario.
   cmdSalir.Left = Width - 820
   fraBuscar.Width = cmdSalir.Left - fraBuscar.Left - 50
   txtBuscar.Width = fraBuscar.Width - 240
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   porstMain.Close
   pocnnMain.Close
   Set porstMain = Nothing
   Set porstCoCfg = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdAceptar_Click()
   
  With porstSGPms
    If .RecordCount <> 0 Then .MoveFirst
      .Find "cLlave='" & gsCodUsr & porstMain!codemp & gsCodSis & "'"
      If .EOF Then
        MsgBox Choose(gsIdioma, "No tiene acceso a la empresa seleccionada.", "You don't have access to the selected company"), vbInformation
        Exit Sub
      End If
  End With

  ' Validar si tien peroiodo habilitado
  If Not ValidaEjercicio(porstMain!codemp, psAnoAct) Then
    MsgBox Choose(gsIdioma, "Periodo seleccionado no se encuentra habilitado; Verificar", "Selected period is qualified; Verify"), vbInformation
    Exit Sub
  End If

  gsAnoAct = psAnoAct
  gsMesAct = Format(cmbPeriodo.ListIndex, "00")
  gsCodEmp = porstMain!codemp
  gsRazEmp = porstMain!RazEmp
  gsRUCEmp = IIf(IsNull(porstMain!RUCEmp), "", porstMain!RUCEmp)
  gsDirEmp = IIf(IsNull(porstMain!direccion), "", porstMain!direccion)
  gsLocEmp = IIf(IsNull(porstMain!localidademp), "", porstMain!localidademp)
  gsGirEmp = IIf(IsNull(porstMain!actividademp), "", porstMain!actividademp)
  gsRepEmp = IIf(IsNull(porstMain!repnombre), "", porstMain!repnombre & ", ") & IIf(IsNull(porstMain!repapepaterno), "", porstMain!repapepaterno & " ") & IIf(IsNull(porstMain!repapematerno), "", porstMain!repapematerno)
  gsRepDNIEmp = IIf(IsNull(porstMain!repdocumento), "", porstMain!repdocumento)
  gsConEmp = IIf(IsNull(porstMain!connombre), "", porstMain!connombre & ", ") & IIf(IsNull(porstMain!conapepaterno), "", porstMain!conapepaterno & " ") & IIf(IsNull(porstMain!conapematerno), "", porstMain!conapematerno)
  gsConDNIEmp = IIf(IsNull(porstMain!condocumento), "", porstMain!condocumento)
  
  gsNomBDS = "C" & gsCodEmp & gsAnoAct
  gsNomBDS = "sysmacon"
   
  frmMain.lblVar(0) = porstMain!RazEmp
  frmMain.lblVar(2) = gsAnoAct
  frmMain.lblVar(3) = gsMesAct

 '[Configuración de la aplicación.
  With porstCoCfg
    .ActiveConnection = CONNSTRG & gsNomBDS
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
    .Open
    gsTpoMon_Fnc = !TpoMon_Fnc
    gnIndMNE = !IndMNE
    gsNivCta = "2" & IIf(!CodCta_Nv3, "3", "") & IIf(!CodCta_Nv4, "4", "") & IIf(!CodCta_Nv5, "5", "") & IIf(!CodCta_Nv6, "6", "") & IIf(!CodCta_Nv7, "7", "") & IIf(!CodCta_Nv8, "8", "")
    gsNivCCo = "2" & IIf(!CodCCo_Nv3, "3", "") & IIf(!CodCCo_Nv5, "5", "")
    gsCodTDc_Pcp = IIf(IsNull(!COdTDC_Pcp), "", !COdTDC_Pcp)
    gsCodTDc_Rtc = IIf(IsNull(!CodTDc_Rtc), "", !CodTDc_Rtc)
    gsCodCta_Pcp = IIf(IsNull(!COdCta_Pcp), "", !COdCta_Pcp)
    gsCodCta_Rtc = IIf(IsNull(!CodCta_Rtc), "", !CodCta_Rtc)
    gsIndRtc = IIf(IsNull(!IndRtc), "N", !IndRtc)
    gsIndPcp = IIf(IsNull(!IndPcp), "N", !IndPcp)
    gsTpoGlo_Rtc = IIf(IsNull(!TpoGlo_Rtc), "0", !TpoGlo_Rtc)
    gsGloDoc_Rtc(0) = ""
    gsGloDoc_Rtc(1) = IIf(IsNull(!GloDocr_Rtc), "", !GloDocr_Rtc)
    gsGloDoc_Rtc(2) = IIf(IsNull(!GloDocn_Rtc), "", !GloDocn_Rtc)
    gnProDestino = IIf(IsNull(!prodestino), 0, !prodestino)
    
    ' Inicializo apertura y cierre
    gnFrances = !ejerfran
    gsMesApe = IIf(!ejerfran, "09", "01")
    gsMesCie = IIf(!ejerfran, "08", "12")
    
    gsCodDro_Ing = IIf(IsNull(!coddro_ing), "", !coddro_ing)
    gsCodDro_Egr = IIf(IsNull(!coddro_egr), "", !coddro_egr)
    .Close
    
    ' Parametros adicionales de configuración
    .Source = "SELECT PctIGV, PctIGV1, PctIGV2, PctISC, PctIR4, PctIES, PctRtc, PctPcp, ImpUIT "
    .Source = .Source & "FROM TGCfg "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
    .Open
    gnPctIGV = CDec(!PctIGV)
    gnPctIGV1 = CDec(!PctIGV1)
    gnPctIGV2 = CDec(!PctIGV2)
    gnPctISC = CDec(!PctISC)
    gnPctIR4 = CDec(!PctIR4)
    gnPctIES = CDec(!PctIES)
    gnPctRtc = CDec(!PctRtc)
    gnPctPcp = CDec(!PctPcp)
    gnImpUIT = CDec(!ImpUIT)
    .Close
  End With
  '[Propio del Proyecto.
  gpCamposSaldos
  gpCieMes
  ']
  Unload Me
End Sub

Public Sub cmdRefrescar_Click()
   gpTUg_Refrescar Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = " ORDER BY "
'   Select Case pnColumnaOrd            'Cambiar.
'   Case 1, 2, 3
'      psConnStrgOrde = psConnStrgOrde & "1, 2, 3"
'   Case Else
      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
'   End Select
   With porstMain
      .Close
      .Source = psConnStrgSele & psConnStrgOrde
      .Open
   End With
   Set dgrMain.DataSource = porstMain
   ppDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyHome
      porstMain.MoveFirst
   Case vbKeyEnd
      porstMain.MoveLast
   End Select
End Sub

Private Sub lstAnoAct_Click()
   psAnoAct = lstAnoAct.Text
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With porstMain
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
      porstMain.Bookmark = dvRegistroActual
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
            .Item(dnNum).Width = 600
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Razon Social", "Firm Name")
            .Item(dnNum).Width = 3800
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
'   cmdNuevo.Enabled = taOpciones(0)
'   cmdEliminar.Enabled = taOpciones(1)
'   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

