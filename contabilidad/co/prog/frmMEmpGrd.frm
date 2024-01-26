VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMEmpGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   5055
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9630
      _ExtentX        =   16986
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
      ScaleWidth      =   9630
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   9630
      Begin VB.CommandButton cmdActualiza 
         BackColor       =   &H80000013&
         Caption         =   "&Slq"
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
         Left            =   4230
         Picture         =   "frmMEmpGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   560
      End
      Begin VB.CommandButton CmdNueAno 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo &Año"
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
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   560
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
         Picture         =   "frmMEmpGrd.frx":0672
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
         Left            =   6120
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
         Left            =   8955
         Picture         =   "frmMEmpGrd.frx":07BC
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
         Picture         =   "frmMEmpGrd.frx":0906
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
         Picture         =   "frmMEmpGrd.frx":0A08
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
         Picture         =   "frmMEmpGrd.frx":0B0A
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
         Picture         =   "frmMEmpGrd.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
      Begin MSComctlLib.Toolbar toolbar 
         Height          =   570
         Left            =   4920
         TabIndex        =   12
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1005
         ButtonWidth     =   1296
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exportar"
               Object.ToolTipText     =   "Exportar Registro de Documentos a Excel"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A1"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A2"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A3"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A4"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   1080
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMEmpGrd.frx":0D0E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMEmpGrd.frx":0E68
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMEmpGrd.frx":0FC2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMEmpGrd.frx":1384
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMEmpGrd.frx":1A4E
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmMEmpGrd"
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

Private Sub cmdActualiza_Click()
frmOQuery.Show vbModal
End Sub

Private Sub CmdNueAno_Click()
   gsCodEmp = Me.dgrMain.Columns(0)
  frmMEmpAno.Show vbModal
End Sub



'[Propio del formulario.
'Public uorstTGDtt As ADODB.Recordset
']

Private Sub Form_Load()
'ini 2015-11-19 rpt consoli x empresa
toolbar.Buttons(1).ButtonMenus(1).Text = "Compras al Mes"
toolbar.Buttons(1).ButtonMenus(2).Text = "Ventas al Mes"
toolbar.Buttons(1).ButtonMenus(3).Text = "Balance de comprobación al mes"
toolbar.Buttons(1).ButtonMenus(4).Text = "Balance de comprobación del mes" '2015-12-01 balance del mes x cliente
'fin 2015-11-19 rpt consoli x empresa

 '[Recordsets                          'Cambiar.
   
   psConnStrgSele = "SELECT CodEmp , RazEmp , RUCEmp, direccion, localidademp, actividademp, "
   psConnStrgSele = psConnStrgSele & "repapepaterno, repapematerno, repnombre, repdocumento, "
   psConnStrgSele = psConnStrgSele & "conapepaterno, conapematerno, connombre, condocumento, "
   psConnStrgSele = psConnStrgSele & "usrcre, fyhcre, usrmdf, fyhmdf "
   '2016-06-02 adicion campo EstAct en Empresa psConnStrgSele = psConnStrgSele & ",buencontri " '2015-08-27 ctr obligac sunat
   psConnStrgSele = psConnStrgSele & ",buencontri,EstEmp "
'ini 2015-01-07 adiciono imagen empresa
'   psConnStrgSele = psConnStrgSele & ",logoemp "
'fin 2015-01-07 adiciono imagen empresa
   psConnStrgSele = psConnStrgSele & "FROM TGEmp "
   psConnStrgOrde = "ORDER BY 1"

   Set uocnnMain = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
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
      .Properties("Unique Table").Value = "TGEmp"
   End With
 ']
   
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  cmdNuevo.Caption = Choose(gsIdioma, "Nuevo &Año", "New &Year")
  cmdActualiza.Caption = Choose(gsIdioma, "&SQL", "&SQL")
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

If gsNvlUsr = 0 Then
   gpTUg_Nuevo Me, frmMEmp             'Cambiar Formulario de Datos.
Else
   MsgBox Choose(gsIdioma, "Acceso no Permitido", "Access not allowed"), vbCritical
   Exit Sub
End If

End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err
   
   If uorstMain.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   With frmMEmp                        'Cambiar Formulario de Datos.
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

Public Sub cmdImprimir_Click(Index As Integer)
 '[Datos del formulario de impresión.  'Cambiar.
   frmLEmp.Caption = Choose(gsIdioma, "Listado de Empresas", "Listing of Enterprises")
   frmLEmp.Show vbModal
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

'ini 2015-11-19 rpt consoli x empresa
Private Sub toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case ButtonMenu.Key
   Case "A1": pExportaRegCmpr 2
   Case "A2": pExportaRegVta 2
'ini 2015-12-01 balance del mes x cliente
   'Case "A3": pExportaBalCom 2
   Case "A3": pExportaBalCom2 1 ' al mes
   Case "A4": pExportaBalCom2 2 ' del mes
'fin 2015-12-01 balance del mes x cliente
  End Select

End Sub
'fin 2015-11-19 rpt consoli x empresa

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
            .Item(dnNum).Width = 600
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Razon Social", "Firm Name")
            .Item(dnNum).Width = 6200
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "R.U.C", "R.U.T.")
            .Item(dnNum).Width = 1100
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

'ini 2015-11-19 rpt consoli x empresa
Private Sub pExportaRegCmpr(TpoRpt As Integer)

    Dim oProgress As New frmzProgressBar
    oProgress.Show
    oProgress.pgbProgreso.Value = 0: oProgress.pgbProgreso.Min = 0
    oProgress.pgbProgreso.Max = 5
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Min
    oProgress.Caption = "Procesando Compras"

'viene de reportes/Registros/Registros de compras
'se le quita la relacion con diario saldo pagados
'2015-11-19 rpt consoli x empresaPrivate Sub pExporta(TpoRpt As Integer)

'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err

'ini 2015-11-19 rpt consoli x empresa
    Dim pocnnMain As ADODB.Connection
    Set pocnnMain = New ADODB.Connection
    With pocnnMain
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
'fin 2015-11-19 rpt consoli x empresa



    Dim pocnnTmp As ADODB.Connection '2014-04-14 Query timeout expired
    Set pocnnTmp = New ADODB.Connection '2014-04-14 Query timeout expired
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
    
        
    Dim cCadReporte  As String
    Dim sTabla As String
    sTabla = "xlsCprCab"
   'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1

        cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
        cCadReporte = cCadReporte & "SELECT"
        cCadReporte = cCadReporte & "    a.CodEmp, e.razemp,"
        cCadReporte = cCadReporte & "    CONCAT(a.pdoano,a.mespvs,'00') as CPERIODO,"
        cCadReporte = cCadReporte & "    concat(a.CodDro,a.NroCpb) as CNUMREGOPE,"
        cCadReporte = cCadReporte & "    date_format(a.FeEDoc,'%d/%m/%Y')as CFECCOM,"
        cCadReporte = cCadReporte & "    date_format(a.FevDOC,'%d/%m/%Y')as CFECVENPAG,"
        cCadReporte = cCadReporte & "    b.CodTDc AS CTIPDOCCOM,"
        '#IF(b.CodTDc<>'50',a.serdoc,mid(a.serdoc,2,3)) AS CNUMSER,
        cCadReporte = cCadReporte & "    IFNULL(case b.codtdc when '50' then a.codaduana when '52' then a.codaduana when '53' then a.codaduana else a.serdoc end, '-') AS CNUMSER,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.annodua,''),a.annodua,'0') as CEMIDUADSI,"
        '#a.NroDoc AS CNUMDCODFV,
        cCadReporte = cCadReporte & "    IFNULL(CASE b.codtdc when '50' then a.nrodua when '52' then a.nrodua when '53' then a.nrodua else a.nrodoc END, '') AS CNUMDCODFV,"
        cCadReporte = cCadReporte & "    '0' AS COSDCREFIS, MID(c.tpodci,2,1) AS CTIPDIDPRO,c.codaux AS CNUMDIDPRO,"
        
        cCadReporte = cCadReporte & "    replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        cCadReporte = cCadReporte & "    ifnull(MID(c.RazAux,1,60)  ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
        cCadReporte = cCadReporte & "    as CNOMRSOPRO,"
        
        '#IF((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1))<>0.00,(a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),'0.00') AS CBASIMPGRA
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS CBASIMPGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS CIGVGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CBASIMPGNG,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIGVGRANGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CBASIMPSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_ONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CIGVSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CIMPTOTNGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CISC,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS COTRTRICGO,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIMPTOTCOM,"
        cCadReporte = cCadReporte & "    format(a.imptcb,3) * 1  AS CTIPCAM,"
        cCadReporte = cCadReporte & "    IF(ifnull(codtdc_ref,''),date_format(feedoc_ref,'%d/%m/%Y'),'01/01/0001') as CFECCOMMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.codtdc_ref,''),a.codtdc_ref,'00')as CTIPCOMMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.serdoc_ref,''),a.serdoc_ref,'-') as CNUMSERMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.nrodoc_ref,''),a.nrodoc_ref,'-') as CNUMCOMMOD,"
        
        cCadReporte = cCadReporte & "    CASE WHEN a.codtdc_ref='91' THEN Concat(a.serdoc_ref, '-', a.nrodoc_ref) ELSE '-' END as CCOMNODOMI,"
        
        cCadReporte = cCadReporte & "    IF(ifnull(a.NroCDt,''),date_format(a.FehCDt,'%d/%m/%Y'),'01/01/0001') as CEMIDEPDET,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.NroCDt,''),a.NroCDt,'0')    as CNUMDEPDET,"
        
        cCadReporte = cCadReporte & "    ifnull(a.INDRETEN,'')   AS CCOMPGRET,"
        
        cCadReporte = cCadReporte & "    if(MONTH(a.FeEDoc)=mespvs,'1','6') as CESTOPE,"
        
        cCadReporte = cCadReporte & "    '0.00' AS CVALFACIMP ,"
        
        cCadReporte = cCadReporte & "    '' AS CINTDIAMAY,"
        cCadReporte = cCadReporte & "    '' AS CINTKARDEX,"
        cCadReporte = cCadReporte & "    '' AS CINTREG, "
        cCadReporte = cCadReporte & "    tsadetrac "
        cCadReporte = cCadReporte & "    ,'' xCol1 " '2015-05-14
        cCadReporte = cCadReporte & "    ,'' xCol2 " '2015-05-14
        cCadReporte = cCadReporte & "    ,a.tpomon " '2015-05-14
        
        cCadReporte = cCadReporte & "    ,replace(format((a.ImpTot_ME * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIMPTOTMEX "  '2015-07-13 adici vta/cpr me
        
        cCadReporte = cCadReporte & "    ,GloDoc " '2015-06-04 adicion glodoc
    
        cCadReporte = cCadReporte & "    ,a.codaux, a.codtdc, a.serdoc, a.NroDoc " '2015-07-15 adicion pgo segun diario

        
        cCadReporte = cCadReporte & "FROM ((((COCprDoc a "
        cCadReporte = cCadReporte & "LEFT JOIN TGTDc b on a.codemp=b.codemp and b.CodTDc = a.CodTDc) "
        cCadReporte = cCadReporte & "LEFT JOIN TGAux c on a.codemp=c.codemp and c.CodAux = a.CodAux) "
        cCadReporte = cCadReporte & "LEFT JOIN CODro d ON a.codemp=d.codemp and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
        cCadReporte = cCadReporte & "LEFT JOIN " & gsNomBDC & ".tgemp e ON  a.codemp=e.codemp)  "
        cCadReporte = cCadReporte & "WHERE "
        'fin 2015-11-19 rpt consoli x empresa cCadReporte = cCadReporte & "    a.codemp='" & gsCodEmp & "' AND "
        
        'ini 2015-03-24
'        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
'        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & Left(cmbEjercicio.Text, 2) & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        If TpoRpt = 1 Then
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' "
'''        Else
        ElseIf TpoRpt = 2 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
             cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        Else
'            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
             cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        End If
'ini 2016-06-02/06 filtra reportes por EstAct
      cCadReporte = cCadReporte & " AND e.estemp='" & ESTEMPR_ACT & "' "
'fin 2016-06-02/06 filtra reportes por EstAct
        
        'fin 2015-03-24
        'cCadReporte = cCadReporte & "    a.pdoano='" & sPdoAnoFin & "' AND "
        '2015-03-23  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & gsMesAct & "' AND "
        '2015-03-20 cambio de periodo cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & "201502" & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        '2015-03-23 cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        
'ini 2015-11-19 rpt consoli x empresa

''ini 2015-01-09 adiciona ruc
'      If Trim(txtDato0_Text) <> "" Then
'          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato0_Text) & "' "
'        End If
''fin 2015-01-09 adiciona ruc

'fin 2015-11-19 rpt consoli x empresa
        
        cCadReporte = cCadReporte & "ORDER BY a.CodEmp,a.pdoano, a.mespvs ,a.CodDro, a.NroCpb ASC "
        '2015-11-19 rpt consoli x empresa cCadReporte = cCadReporte & "ORDER BY a.pdoano, a.mespvs ,a.CodDro, a.NroCpb ASC "

    
    pocnnTmp.Execute cCadReporte
    
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
        
'ini 2015-07-15 adicion pgo segun diario
    sTabla = "tmp_xls_pdte"
    '2015-11-19 rpt consoli x empresa  pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
'saldos de documento segun reporte historico
    cCadReporte = cCadReporte & "SELECT "
    cCadReporte = cCadReporte & "    a.pdoano AS cAno, a.MesPvs, a.CodCta, a.CodAux,a.CodTDc,"
    cCadReporte = cCadReporte & "    a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, Null AS codcco,"
    'cCadReporte = cCadReporte & "    Null AS detcco, CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocum, a.FehOpe, a.FeEDoc, a.FeVDoc,"
    cCadReporte = cCadReporte & "    Null AS detcco,"
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocum, "
    cCadReporte = cCadReporte & "a.FehOpe, a.FeEDoc, a.FeVDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, b.RazAux, "
    
    'cCadReporte = cCadReporte & "    a.RefDoc, a.GloIte AS GloIte, b.RazAux, (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
    cCadReporte = cCadReporte & "(CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) AS cDebeMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) AS cHaberMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END) AS cDebeME, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) AS cHaberME "
    
'    cCadReporte = cCadReporte & "    (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END) AS cDebeMN, (CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END) AS cHaberMN,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END) AS cDebeME, (CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END) AS cHaberME"
    cCadReporte = cCadReporte & "    ,a.TpoPvs"
    cCadReporte = cCadReporte & "    ,CONCAT(year(a.FehOpe),'-',LPAD(month(a.FehOpe),2,'0'),'-',LPAD(day(a.FehOpe),2,'0'),'-',a.CodDro,'-',a.NroCpb,'-',a.Nroite) AS x_clave "
    cCadReporte = cCadReporte & "FROM ((((COCpbDet a "
    cCadReporte = cCadReporte & "    LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    cCadReporte = cCadReporte & "    LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    cCadReporte = cCadReporte & "    LEFT JOIN Cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.Codcta=d.Codcta) "
    cCadReporte = cCadReporte & "    LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.codcco=e.codcco) "
'    cCadReporte = cCadReporte & "WHERE a.codemp='010' "
'    cCadReporte = cCadReporte & "    AND a.pdoano='2014' "
    cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "    AND a.pdoano='" & gsAnoAct & "' "
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta, 2)>='01' AND LEFT(a.codcta, 2)<='FF' "
    cCadReporte = cCadReporte & "    AND (a.ImpMN<> 0.00 OR a.ImpME<> 0.00) "
    'cCadReporte = cCadReporte & "    AND a.Mespvs <='03' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND a.Mespvs <='" & gsMesAct & "' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND IFNULL(a.NroDoc, '') <>'' AND d.inddoc='1' "
    'cCadReporte = cCadReporte & "    # AND a.CodAux='10097267265'"
    cCadReporte = cCadReporte & "    AND a.TpoPvs='" & TPOPVS_CAN & "' " 'TPOPVS_CAN
    'cCadReporte = cCadReporte & "    AND a.TpoPvs='C' " 'TPOPVS_CAN
'2016-03-14 ini erro cuenta 422,428 se mezclan
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta,3) <> '422' " '2015-09-03 erro cuenta anticipo duplicado
    cCadReporte = cCadReporte & "    AND LEFT(a.codcta,3) = '421' "
'2016-03-14 fin erro cuenta 422,428 se mezclan
   
'ini 2015-11-19 rpt consoli x empresa

''ini 2015-01-09 adiciona ruc
'      If Trim(txtDato0_Text) <> "" Then
'          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato0_Text) & "' "
'        End If
''fin 2015-01-09 adiciona ruc

'fin 2015-11-19 rpt consoli x empresa
    
    cCadReporte = cCadReporte & "ORDER BY a.codcta, a.codaux, a.codtdc, a.serdoc, a.NroDoc, a.TpoPvs, a.MesPvs, a.FehOpe "
    
    '2015-11-19 rpt consoli x empresa  pocnnTmp.Execute cCadReporte

'*********************
    sTabla = "tmp_xls_pdte2"
    '2015-11-19 rpt consoli x empresa  pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "    codcta,codaux,cdocum,min(x_clave) x_clave "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
cCadReporte = cCadReporte & "GROUP BY codcta, codaux,cdocum # x_clave "
cCadReporte = cCadReporte & "ORDER BY codcta, codaux,cdocum #x_clave "

    '2015-11-19 rpt consoli x empresa  pocnnTmp.Execute cCadReporte

'*********************
    sTabla = "tmp_xls_pdte3"
    '2015-11-19 rpt consoli x empresa  pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")

cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "* "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
cCadReporte = cCadReporte & "Where x_clave "
cCadReporte = cCadReporte & "    IN (select x_clave from tmp_xls_pdte2) "
    '2015-11-19 rpt consoli x empresa pocnnTmp.Execute cCadReporte
    
'*********************
    
'fin 2015-07-15 adicion pgo segun diario
    
    
'ini exporta datos a excel

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       
'       .Source = "SELECT * FROM " & ps_Prefijo & sTabla

'ini 2015-07-15 adicion pgo segun diario
.Source = "SELECT "
.Source = .Source & "    CodEmp, razemp, "
.Source = .Source & "    CPERIODO,CNUMREGOPE,CFECCOM,CFECVENPAG,CTIPDOCCOM,"
.Source = .Source & "    CNUMSER,CEMIDUADSI,CNUMDCODFV,COSDCREFIS,CTIPDIDPRO,"
.Source = .Source & "    CNUMDIDPRO,CNOMRSOPRO,CBASIMPGRA,CIGVGRA,CBASIMPGNG,"
.Source = .Source & "    CIGVGRANGV,CBASIMPSCF,CIGVSCF,CIMPTOTNGV,CISC,"
.Source = .Source & "    COTRTRICGO,CIMPTOTCOM,CTIPCAM,CFECCOMMOD,CTIPCOMMOD,"
.Source = .Source & "    CNUMSERMOD,CNUMCOMMOD,CCOMNODOMI,CEMIDEPDET,CNUMDEPDET,"
.Source = .Source & "    CCOMPGRET,CESTOPE,CVALFACIMP,CINTDIAMAY,CINTKARDEX,"
.Source = .Source & "    CINTREG,tsadetrac,xCol1,xCol2,tpomon,"

.Source = .Source & "    CIMPTOTMEX,GloDoc "
'ini 2015-11-19 rpt consoli x empresa
''#   a.*,,b.FehOpe
'.Source = .Source & "    b.FehOpe,"
'.Source = .Source & "    IFNULL(b.cDebeMN,0)-IFNULL(b.cHaberMN,0) PgoMN,"
'.Source = .Source & "    IFNULL(b.cDebeME,0)-IFNULL(b.cHaberME,0) PgoME "
'fin 2015-11-19 rpt consoli x empresa

.Source = .Source & "FROM xlsCprCab a "

'ini 2015-11-19 rpt consoli x empresa
'.Source = .Source & "LEFT JOIN tmp_xls_pdte3 b "
'.Source = .Source & "    ON a.CodAux=b.CodAux and a.CTIPDOCCOM=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc "
'fin 2015-11-19 rpt consoli x empresa

'fin 2015-07-15 adicion pgo segun diario
       
       .Open
    End With
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1

    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
 
    'Set oSheet = oWBook.Worksheets(1)
 

    '*Set oExcel = New Excel.Application
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

    Set oWBook = oExcel.Workbooks.Add
    '*Set oWBook = oExcel.Workbooks.Open(dlbDirectorio(0).path & xArchPeriodo, , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    '*Set oSheet = oWBook.Worksheets("Clientes")
     Set oSheet = oWBook.Worksheets(1)
    '*oExcel.Visible = True

    With oSheet
    
'        oSheet.Select
'        Columns("M:V").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"

        oSheet.Select
        
        .Cells(1, 1).Value = "Registro de Compras"
        
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = "Registro de Compras"
        nRowI = nRowI + 2
        Dim x1 As Integer
        Dim x As Integer
        x = x + 1: .Cells(nRowI, x).Value = "Empresa"
        x = x + 1: .Cells(nRowI, x).Value = "Detalle"
        x = x + 1: .Cells(nRowI, x).Value = "Periodo"
        x = x + 1: .Cells(nRowI, x).Value = "Nº Reg."
        x = x + 1: .Cells(nRowI, x).Value = "F.Cmpra"
        x = x + 1: .Cells(nRowI, x).Value = "F. Pago"
        x = x + 1: .Cells(nRowI, x).Value = "T.Doc"
        x = x + 1: .Cells(nRowI, x).Value = "Serie"
        x = x + 1: .Cells(nRowI, x).Value = "CemiDuadsi"
        x = x + 1: .Cells(nRowI, x).Value = "Nº Doc."
        x = x + 1: .Cells(nRowI, x).Value = "COSDCREFIS"
        x = x + 1: .Cells(nRowI, x).Value = "T.Prv"
'10
        x = x + 1: .Cells(nRowI, x).Value = "RUC"
        x = x + 1: .Cells(nRowI, x).Value = "R.Social"
        x = x + 1: .Cells(nRowI, x).Value = "B. Gravada"
        x = x + 1: .Cells(nRowI, x).Value = "IGV Grab"
        x = x + 1: .Cells(nRowI, x).Value = "B. G/N Gr"
        x = x + 1: .Cells(nRowI, x).Value = "IGV G/N Gr"
        x = x + 1: .Cells(nRowI, x).Value = "B. Sin CF"
        x = x + 1: .Cells(nRowI, x).Value = "Igv S CF"
        x = x + 1: .Cells(nRowI, x).Value = "CIMPTOTNGV"
        x = x + 1: .Cells(nRowI, x).Value = "CISSC"
'20
        x = x + 1: .Cells(nRowI, x).Value = "COTRTRICGO"
        x = x + 1: .Cells(nRowI, x).Value = "CIMPTOTCOM"
        x = x + 1: .Cells(nRowI, x).Value = "CTIPCAM"
        x = x + 1: .Cells(nRowI, x).Value = "CFECCOMMOD"
        x = x + 1: .Cells(nRowI, x).Value = "CTIPCOMMOD"
        x = x + 1: .Cells(nRowI, x).Value = "CNUMSERMOD"
        x = x + 1: .Cells(nRowI, x).Value = "CNUMCOMMOD"
        x = x + 1: .Cells(nRowI, x).Value = "CCOMNODOMI"
        x = x + 1: .Cells(nRowI, x).Value = "CEMIDEPDET"
        x = x + 1: .Cells(nRowI, x).Value = "CNUMDEPDET"
'30
        x = x + 1: .Cells(nRowI, x).Value = "CCOMPGRET"
        x = x + 1: .Cells(nRowI, x).Value = "CESTOPE"
        x = x + 1: .Cells(nRowI, x).Value = "CVALFACIMP"
        x = x + 1: .Cells(nRowI, x).Value = "CINTDIAMAY"
        x = x + 1: .Cells(nRowI, x).Value = "CINTKARDEX"
        x = x + 1: .Cells(nRowI, x).Value = "CINTREG"
        x = x + 1: .Cells(nRowI, x).Value = "tsadetrac"
        x = x + 1: .Cells(nRowI, x).Value = "DetaDetrac"
        x = x + 1: .Cells(nRowI, x).Value = "PorcDetra"
        x = x + 1: .Cells(nRowI, x).Value = "TpoMon"
'40
        x = x + 1: .Cells(nRowI, x).Value = "Total ME"
        x = x + 1: .Cells(nRowI, x).Value = "Glosa"
        
'ini 2015-11-19 rpt consoli x empresa
'        x = x + 1: .Cells(nRowI, x).Value = "F.Pago"
'        x = x + 1: .Cells(nRowI, x).Value = "1er Pgo MN"
'        x = x + 1: .Cells(nRowI, x).Value = "1er Pgo ME"
'fin 2015-11-19 rpt consoli x empresa
     
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        .Columns.AutoFit ' ajusta el ancho de las columnas
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
        'Sheets(oSheet).Select
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'solo sale error en esta        Columns("M:M").Select
'        Selection.NumberFormat = "#,##0.00"

'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("O:O").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("P:P").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("Q:Q").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("R:R").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("S:S").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("T:T").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("U:U").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("V:V").Select
'        Selection.NumberFormat = "#,##0.00"
        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'        oSheet.Select
'        Cells(1, 1).Select

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 2).Value)
'        Next


 'ini 2015-07-02 adic tabla detrac
'*********************************
        Dim uorstcodetrac As ADODB.Recordset
        Set uorstcodetrac = New ADODB.Recordset
        Set uorstcodetrac = fRstDetrac(pocnnMain, uorstcodetrac)
        'Set uorstcodetrac = fRstDetrac(uocnnMain, uorstcodetrac)
'        With uorstCoDetrac
'           .ActiveConnection = pocnnMain
'           .Source = "SELECT coddetrac, " & Choose(gsIdioma, "detdetrac", "detdetracx") & " AS DetDetrac,tsadetrac ,  "
'           .Source = .Source & "codemp "
'           .Source = .Source & "FROM codetrac  "
'           .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
'           .Source = .Source & "AND estdetrac ='" & ESTDETRAC_ACT & "' "
'           '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'           '.Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
'           .CursorType = adOpenDynamic
'           .LockType = adLockOptimistic
'           .Open
'        End With
        
        
       xrow1 = nRowI
        Dim nContador As Integer
        Dim s_Contenido As String
        Dim n_Detraccion As Double
        Dim s_detalle As String
        Do While Len(Trim(.Cells(xrow1, 2).Value)) <> 0
            s_Contenido = Left(.Cells(xrow1, 37).Value, 5)
            With uorstcodetrac
                If .RecordCount > 0 Then .MoveFirst
                    .Find "coddetrac='" & s_Contenido & "'"
                    If Not .EOF Then
                        oSheet.Cells(xrow1, 38).Value = !coddetrac
                        '2015-07-08 cambio de decima a % oSheet.Cells(xrow1, 39).Value = !pctdetrac * 100
                        oSheet.Cells(xrow1, 39).Value = !pctdetrac
                   End If
            End With
            xrow1 = xrow1 + 1
        Loop
        
        uorstcodetrac.Close
        Set uorstcodetrac = Nothing


'*********************************
''       xrow1 = nRowI
''        Dim nContador As Integer
''        Dim s_Contenido As String
''        Dim n_Detraccion As Double
''        Dim s_detalle As String
''        Do While Len(Trim(.Cells(xrow1, 2).Value)) <> 0
''            'MsgBox (.Cells(xrow1, 37).Value)
''            's_Contenido = Left(.Cells(xrow1, 37).Value, 3)
''            s_Contenido = Left(.Cells(xrow1, 37).Value, 5)
''            'ini 2014-04-05 reclasificacion de cod detraccion
''            For nContador = 1 To UBound(aDtraccDet, 1)
''            'If Left(aDtraccDet(nContador), 3) = s_Contenido Then
''            If Left(aDtraccDet(nContador), 5) = s_Contenido Then
''                n_Detraccion = aDtraccPor(nContador)
''                s_detalle = aDtraccDet(nContador)
''                s_detalle = Mid(s_detalle, 7)
''                .Cells(xrow1, 38).Value = s_detalle
''                .Cells(xrow1, 39).Value = n_Detraccion * 100
''               Exit For
''            End If
''            Next nContador
''            xrow1 = xrow1 + 1
''        Loop
'fin 2015-07-02 adic tabla detrac
        
    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing


'fin exporta datos a excel
  pocnnMain.Execute fDropTable("xlsCprCab", 1)
  
'ini 2015-11-19 rpt consoli x empresa
  pocnnMain.Execute fDropTable("tmp_xls_pdte", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte2", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte3", 1)
'fin 2015-11-19 rpt consoli x empresa

'fDropTable
   porstTmp.Close
   pocnnTmp.Close
   Set porstTmp = Nothing
   Set pocnnTmp = Nothing

'ini 2015-11-19 rpt consoli x empresa
   pocnnMain.Close
   Set pocnnMain = Nothing
'fin 2015-11-19 rpt consoli x empresa

    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
    Unload oProgress          ' Unload progress bar window

  Exit Sub
Err:
    MsgBox (TEXT_6001)
  If pocnnTmp.State = adStateOpen Then
    porstTmp.Close
    pocnnTmp.Close
    Set porstTmp = Nothing
    Set pocnnTmp = Nothing
    
'ini 2015-11-19 rpt consoli x empresa
   pocnnMain.Close
   Set pocnnMain = Nothing
'fin 2015-11-19 rpt consoli x empresa
    
  End If
End Sub

'****************************************************************************
'Private Sub pExporta(TpoRpt As Integer)
'viene de reportes/Registros/Registros de compras
'se le quita la relacion con diario saldo pagados

Private Sub pExportaRegVta(TpoRpt As Integer)
    Dim oProgress As New frmzProgressBar
    oProgress.Show
    oProgress.pgbProgreso.Value = 0: oProgress.pgbProgreso.Min = 0
    oProgress.pgbProgreso.Max = 5
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Min
    oProgress.Caption = "Procesando Ventas"

'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err

'ini 2015-11-19 rpt consoli x empresa
    Dim pocnnMain As ADODB.Connection
    Set pocnnMain = New ADODB.Connection
    With pocnnMain
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
'fin 2015-11-19 rpt consoli x empresa


    Dim pocnnTmp As ADODB.Connection '2014-04-14 Query timeout expired
    Set pocnnTmp = New ADODB.Connection '2014-04-14 Query timeout expired
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
    
    Dim cCadReporte  As String
    Dim sTabla As String
    sTabla = "xlsVtaCab"
   'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
    

        cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
    cCadReporte = cCadReporte & "SELECT"
    cCadReporte = cCadReporte & "    a.CodEmp, e.razemp," '2015-11-19 rpt consoli x empresa
    cCadReporte = cCadReporte & "    concat(a.pdoano,a.mespvs,'00') AS VPERIODO,"
    cCadReporte = cCadReporte & "    concat(a.CodDro,a.NroCpb) as VNUMREGOPE,"
    cCadReporte = cCadReporte & "    date_format(a.feedoc,'%d/%m/%Y')as VFECCOM,"
    cCadReporte = cCadReporte & "    date_format(a.FevDOC,'%d/%m/%Y')as VFECVENPAG,"
    cCadReporte = cCadReporte & "    b.CodTDc as VTIPDOCCOM, a.SerDoc AS VNUMSER, a. NroDoc AS VNUMDOCCOI,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.NroDoc_Fin,''),a.NroDoc_Fin,'0') AS VNUMDOCCOF,"
    cCadReporte = cCadReporte & "    MID(c.tpodci,2,1) AS VTIPDIDCLI,"
    cCadReporte = cCadReporte & "    c.Codaux AS VNUMDIDCLI,"
    cCadReporte = cCadReporte & "    replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
    cCadReporte = cCadReporte & "    ifnull(MID(c.RazAux,1,60) ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
    cCadReporte = cCadReporte & "    as VAPENOMRSO,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpExp_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VVALFACEXP,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VBASIMPGRA,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTEXO,"
    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTINA,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VISC,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIGVIPM,"
    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VBASIMIVAP,"
    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIVAP,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VOTRTRICGO,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTCOM,"
    cCadReporte = cCadReporte & "    format(a.imptcb,3) * 1 AS VTIPCAM,"
    cCadReporte = cCadReporte & "    IF(ifnull(codtdc_ref,''),date_format(feedoc_ref,'%d/%m/%Y'),'01/01/0001') as VFECCOMMOD,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.codtdc_ref,''),a.codtdc_ref,'00') as VTIPCCOMOD,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.serdoc_ref,''),a.serdoc_ref,'-')  as VNUMSERMOD,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.nrodoc_ref,''),a.nrodoc_ref,'-')  as VNUMCOMMOD,"
    cCadReporte = cCadReporte & "    IF(a.ImpTot_MN <>0.00,'1','2') as VESTOPE,"
    cCadReporte = cCadReporte & "    '' AS VINTDIAMAY,"
    cCadReporte = cCadReporte & "    '' AS VINTKARDEX,"
    cCadReporte = cCadReporte & "    '' AS VINTREG "
    cCadReporte = cCadReporte & "    ,a.TpoMon " '2015-05-14
    
    cCadReporte = cCadReporte & "    ,replace(format((a.ImpTot_ME * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTMEX " '2015-07-13 adici vta/cpr me
    
    cCadReporte = cCadReporte & "    ,GloDoc " '2015-06-04 adicion glodoc
    cCadReporte = cCadReporte & "    ,date_format(a.feedoc,'%d/%m/%Y')as fehcdt,nrocdt,tsadetrac,pctdetrac " '2015-07-03 adicion campo detracc vta
       
    '2015-11-19 rpt consoli x empresa cCadReporte = cCadReporte & "FROM (((COVtaDoc a "
    cCadReporte = cCadReporte & "FROM ((((COVtaDoc a "
    cCadReporte = cCadReporte & "LEFT JOIN TGTDc b ON  a.codemp=b.codemp and a.CodTDc=b.CodTDc) "
    cCadReporte = cCadReporte & "LEFT JOIN TGAux c ON  a.codemp=c.codemp  and a.CodAux=c.CodAux) "
    cCadReporte = cCadReporte & "LEFT JOIN CODro d ON  a.codemp=d.codemp  and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
    cCadReporte = cCadReporte & "LEFT JOIN " & gsNomBDC & ".tgemp e ON  a.codemp=e.codemp) "
    cCadReporte = cCadReporte & "WHERE "
    '2015-11-19 rpt consoli x empresa cCadReporte = cCadReporte & "   a.codemp='" & gsCodEmp & "' and "
        'ini 2015-03-24
'    cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) >='" & gsAnoAct & "01" & "'  and "
'    cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & gsAnoAct & Left(cmbEjercicio.Text, 2) & "' AND "
        If TpoRpt = 1 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' AND "
        'Else '2015-09-03 opc historico
        ElseIf TpoRpt = 2 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND  "
'ini 2015-09-03 opc historico
        Else
            'cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND  "
'fin 2015-09-03 opc historico
        End If
'ini 2016-06-02/06 filtra reportes por EstAct
      cCadReporte = cCadReporte & " e.estemp='" & ESTEMPR_ACT & "' AND "
'fin 2016-06-02/06 filtra reportes por EstAct

        'fin 2015-03-24
    'cCadReporte = cCadReporte & "   a.pdoano='" & sPdoAnoFin & "' and "
    '2015-03-23 cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) >='" & gsAnoAct & gsMesAct & "'  and "
    '*cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & "201412" & "' AND "
    '2015-03-23cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & gsAnoAct & gsMesAct & "' AND "
    cCadReporte = cCadReporte & "   IFNULL(a.CodAux, '')<>'' AND  "
    cCadReporte = cCadReporte & "   IFNULL(a.CodDro, '')<>'' "
'ini 2015-11-19 rpt consoli x empresa

''ini 2015-01-09 adiciona ruc
'      If Trim(txtDato0_Text) <> "" Then
'          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato0_Text) & "' "
'      End If
''fin 2015-01-09 adiciona ruc

'fin 2015-11-19 rpt consoli x empresa
    
    
    '2015-11-19 rpt consoli x empresa cCadReporte = cCadReporte & "ORDER BY a.mespvs ,a.CodTDc, a.SerDoc, a.NroDoc  ASC "
    cCadReporte = cCadReporte & "ORDER BY a.CodEmp,a.mespvs ,a.CodTDc, a.SerDoc, a.NroDoc  ASC "

    
    pocnnTmp.Execute cCadReporte
    
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
    
'ini exporta datos a excel

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       .Source = "SELECT * FROM " & ps_Prefijo & sTabla
       .Open
    End With

        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1

    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
 
    'Set oSheet = oWBook.Worksheets(1)
 

    '*Set oExcel = New Excel.Application
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

    Set oWBook = oExcel.Workbooks.Add
    '*Set oWBook = oExcel.Workbooks.Open(dlbDirectorio(0).path & xArchPeriodo, , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    '*Set oSheet = oWBook.Worksheets("Clientes")
     Set oSheet = oWBook.Worksheets(1)
    '*oExcel.Visible = True

    With oSheet
        oSheet.Select
        
        '.Cells(1, 1).Value = "Registro de Ventas"
        
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = "Registro de Ventas"
        nRowI = nRowI + 2
        Dim x1 As Integer
        Dim x As Integer
        x = 0
        x = x + 1: .Cells(nRowI, x).Value = "Empresa"
        x = x + 1: .Cells(nRowI, x).Value = "Detalle"
        x = x + 1: .Cells(nRowI, x).Value = "Periodo"
        x = x + 1: .Cells(nRowI, x).Value = "Nº Reg."
        x = x + 1: .Cells(nRowI, x).Value = "F.Vta"
        x = x + 1: .Cells(nRowI, x).Value = "F. Pago"
        x = x + 1: .Cells(nRowI, x).Value = "T.Doc"
        x = x + 1: .Cells(nRowI, x).Value = "Serie"
        x = x + 1: .Cells(nRowI, x).Value = "VNUMDOCCCOI"
        x = x + 1: .Cells(nRowI, x).Value = "Nº Doc."
        x = x + 1: .Cells(nRowI, x).Value = "Tpo.Cli"
        x = x + 1: .Cells(nRowI, x).Value = "RUC"
        '10
        x = x + 1: .Cells(nRowI, x).Value = "R.Social"
        x = x + 1: .Cells(nRowI, x).Value = "VVALFACEXP"
        x = x + 1: .Cells(nRowI, x).Value = "VBASIMPGRA"
        x = x + 1: .Cells(nRowI, x).Value = "VIMPTOTEXO"
        x = x + 1: .Cells(nRowI, x).Value = "VIMPTOTINA"
        x = x + 1: .Cells(nRowI, x).Value = "VISC"
        x = x + 1: .Cells(nRowI, x).Value = "VIGVIPM"
        x = x + 1: .Cells(nRowI, x).Value = "VBASIMIVAP"
        x = x + 1: .Cells(nRowI, x).Value = "VIVAP"
        '20
        x = x + 1: .Cells(nRowI, x).Value = "VOTRTRICGO"
        x = x + 1: .Cells(nRowI, x).Value = "CIMPTOTCOM"
        x = x + 1: .Cells(nRowI, x).Value = "VTIPCAM"
        x = x + 1: .Cells(nRowI, x).Value = "VFECCOMMOD"
        x = x + 1: .Cells(nRowI, x).Value = "VTIPCCOMOD"
        x = x + 1: .Cells(nRowI, x).Value = "VNUMSERMOD"
        x = x + 1: .Cells(nRowI, x).Value = "VNUMCOMMOD"
        x = x + 1: .Cells(nRowI, x).Value = "VESTOPE"
        x = x + 1: .Cells(nRowI, x).Value = "VINTDIAMAY"
        x = x + 1: .Cells(nRowI, x).Value = "VINTKARDEX"
        '30
        x = x + 1: .Cells(nRowI, x).Value = "VINTREG"
        x = x + 1: .Cells(nRowI, x).Value = "TpoMon"
        x = x + 1: .Cells(nRowI, x).Value = "Total ME"
        x = x + 1: .Cells(nRowI, x).Value = "Glosa" '2015-06-04 adicion glodoc
'ini 2015-07-03 adicion campo detracc vta
        x = x + 1: .Cells(nRowI, x).Value = "F.Detrac"
        x = x + 1: .Cells(nRowI, x).Value = "Doc.Detrac"
        x = x + 1: .Cells(nRowI, x).Value = "Tsa.Detrac"
        x = x + 1: .Cells(nRowI, x).Value = "% Detrac" '
'fin 2015-07-03 adicion campo detracc vta
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'        Columns("L:L").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("M:M").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("O:O").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("P:P").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("Q:Q").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("R:R").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("S:S").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("T:T").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("U:U").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("V:V").Select
'        Selection.NumberFormat = "#,##0.000"
        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'        oSheet.Select
'        Cells(1, 1).Select

    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing


'fin exporta datos a excel

   porstTmp.Close
   pocnnTmp.Close
   Set porstTmp = Nothing
   Set pocnnTmp = Nothing

'ini 2015-11-19 rpt consoli x empresa
   pocnnMain.Close
   Set pocnnMain = Nothing
'fin 2015-11-19 rpt consoli x empresa

    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
    Unload oProgress          ' Unload progress bar window

  Exit Sub
Err:
    MsgBox (TEXT_6001)
  If pocnnTmp.State = adStateOpen Then
    porstTmp.Close
    pocnnTmp.Close
    Set porstTmp = Nothing
    Set pocnnTmp = Nothing
    
'ini 2015-11-19 rpt consoli x empresa
   pocnnMain.Close
   Set pocnnMain = Nothing
'fin 2015-11-19 rpt consoli x empresa
    
  End If

End Sub
'fin 2015-11-19 rpt consoli x empresa

'ini 2015-11-20 rpt consoli x empresa
'Private Sub cmdImprimir_Click(Index As Integer)
Private Sub pExportaBalCom(TpoRpt As Integer)

'ini 2015-12-01 correcc errore
    Dim oProgress As New frmzProgressBar
    oProgress.Show
    oProgress.pgbProgreso.Value = 0: oProgress.pgbProgreso.Min = 0
    oProgress.pgbProgreso.Max = 5
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Min
    oProgress.Caption = "Procesando Ventas"
'fin 2015-12-01 correcc errore


'pExportaRegVta

'ini 2015-11-19 rpt consoli x empresa
    Dim pocnnMain As ADODB.Connection
    Set pocnnMain = New ADODB.Connection
    With pocnnMain
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
'fin 2015-11-19 rpt consoli x empresa


Dim pnNivCta As Byte
    pnNivCta = 9


  Dim chkRango_Value As Boolean
  chkRango_Value = False 'para que no entre en valor
  
  'rangos de mes
  Dim cmbPeriodo_2_ListIndex As Integer
  Dim cmbPeriodo_3_ListIndex As Integer
  cmbPeriodo_2_ListIndex = 0
  cmbPeriodo_3_ListIndex = 12
  
  Dim cmbPeriodo_0 As String
  Dim cmbPeriodo_1 As String
  cmbPeriodo_0 = "2000"
  cmbPeriodo_1 = "2999"
  
  'moneda
  Dim cboTpoMon_ListIndex As Integer
  cboTpoMon_ListIndex = 0 'solo moneda nacional
  
  Dim optAlcance_0_Value As Boolean
  optAlcance_0_Value = True 'para que no entre en valor
'txtDato0_Text
Dim txtDato0_Text As String
Dim txtDato1_Text As String
txtDato0_Text = "00"
txtDato1_Text = "FF"

Dim chkDivisoria_Value As Boolean
chkDivisoria_Value = False
'********************************************************
  Dim dnContador As Integer, n_Index As Integer
  Dim s_Sentencia As String, s_Sql As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_Moneda As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_MesIni As Integer, n_MesFin As Integer
  Dim l_CreateTB As Boolean
    
  s_AnoIni = Right(IIf(chkRango_Value = vbChecked, cmbPeriodo_0, gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango_Value = vbChecked, cmbPeriodo_1, gsAnoAct), 4)
  
  

  ' Valido el rango de periodos
  If chkRango_Value = vbChecked Then
    s_Mes = Format(cmbPeriodo_2_ListIndex, "00")
    s_Ano = Format(cmbPeriodo_3_ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: 'cmbPeriodo_1.SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: 'cmbPeriodo(3).SetFocus: Exit Sub
  End If
    
  '****ppHabilitacion False
  
  If pnNivCta = 9 Then pnNivCta = Val(Right(gsNivCta, 1))
  s_Moneda = IIf(cboTpoMon_ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
   
 On Error GoTo Err
 
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
  
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore
  
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    n_MesIni = Val(IIf(optAlcance_0_Value, 0, gsMesAct))
    n_MesFin = Val(gsMesAct)
    If chkRango_Value = vbChecked Then
      n_MesIni = Val(IIf(s_Ano = s_AnoIni, cmbPeriodo_2_ListIndex, 1))
      n_MesFin = Val(IIf(s_Ano = s_AnoFin, cmbPeriodo_3_ListIndex, 12))
    End If
    ' Acumulación de saldos
    s_SaldoDeb = "ROUND(": s_SaldoHab = "ROUND("
    For n_Index = n_MesIni To n_MesFin
      s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
    Next n_Index
    s_SaldoDeb = s_SaldoDeb & ", 2)"
    s_SaldoHab = s_SaldoHab & ", 2)"
      
    ' Registros iniciales de saldos
    '2015-11-23 s_Sentencia = "SELECT a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, b.TpoSdo, b.TpoCta, "
    s_Sentencia = "SELECT  a.codemp,CONCAT(a.pdoano,'" & Format(Trim(n_MesFin), "00") & "') pdoano,"
    s_Sentencia = s_Sentencia & " a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, b.TpoSdo, b.TpoCta, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & " AS cSumaD, " & s_SaldoHab & " AS cSumaH, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoDeb & " ELSE 0 END) AS cSumaDt, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoHab & " ELSE 0 END) AS cSumaHt "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCpb ", "")
    s_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    '2015-11-23 s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "WHERE  "
    '2015-11-23 s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato0_Text & "' AND '" & txtDato1_Text & "' "
    If pnNivCta = 2 Then
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    Else
      If chkDivisoria_Value = 1 Then
        s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=2)) "
      Else
        s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
        s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & " AND b.TpoCta=" & TPOCTA_TRA & ")) "
      End If
    End If
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "HAVING (cSumaD + cSumaH + cSumaDt + cSumaHt) > 0 "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY a.CodCta"
    ' Executo la sentencia
    If Not l_CreateTB Then
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trpRngBceCpb ", "")
      's_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS trpRngBceCpb ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trpRngBceCpb "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
   
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore
   
    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
  'With porstMRp
  With porstTmp
'    If .State = adStateOpen Then .Close
'    s_Sentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
'    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
'    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt "
'    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
'    s_Sentencia = s_Sentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
'    s_Sentencia = s_Sentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)) > 0 "
'    s_Sentencia = s_Sentencia & "ORDER BY CodCta"
'    .Source = s_Sentencia
'    .Open
'************************
       .ActiveConnection = pocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
'ini 2015-11-23
    's_Sentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
    s_Sentencia = "SELECT a.codemp,max(e.razemp) razemp,pdoano, "
    s_Sentencia = s_Sentencia & "CodCta, DetCta,  "
'fin 2015-11-23
    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
    '2015-11-23 s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt "
    s_Sentencia = s_Sentencia & "CASE  WHEN ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2) >0 THEN ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2) ELSE 0 END cSumaDt,"
    s_Sentencia = s_Sentencia & "CASE  WHEN ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2) <0 THEN (ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2)) * -1 ELSE 0 END cSumaHt "
    
    '2015-11-23  s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb a "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & gsNomBDC & ".tgemp e ON  a.codemp=e.codemp "
    '2015-11-23 s_Sentencia = s_Sentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
    s_Sentencia = s_Sentencia & "GROUP BY codemp,pdoano,CodCta, DetCta, TpoSdo, TpoCta "
    s_Sentencia = s_Sentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)) > 0 "
    '2015-11-23 s_Sentencia = s_Sentencia & "ORDER BY CodCta"
    s_Sentencia = s_Sentencia & "ORDER BY codemp,pdoano,CodCta"
     .Source = s_Sentencia
      .Open

'***********************
  End With
  
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore

'ini eliminar reporte ***********************************************
'''  s_Sentencia = IIf(chkRango_Value = vbChecked, cmbPeriodo(2).Text & " - " & cmbPeriodo_0.Text, "")
  
'''  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
'''  If usDEstino = PRN_DEST_GRAF Then
'''    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon_ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value, porstMRp
'''    With frmMain.rptMain
'''      '[Datos y parámetros del reporte.  'Cambiar.
'''      .ReportFileName = gsRutRpt & "rptRBceCpb.rpt"
'''      'Fórmular propias.
'''      .Formulas(5) = "mPeriodo='" & s_Sentencia & " " & IIf(optAlcance_0_Value, Choose(gsIdioma, "Acumulado - ", "Accrued - "), "") & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
'''      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
'''      .MarginLeft = unMargenIzquierdo
'''      .WindowState = crptMaximized
'''      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
'''      .Action = 1
'''    End With
'''  Else
'''    Set MRViewer = New MRViewerObject
'''    With MRViewer
'''      .DataRecordSet = porstMRp
'''       .LoadReport gsRutRpt & "rptRBceCpb.mrp"
'''       Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon_ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
'''      '[Parámetros adicionales.
'''       If optAlcance_0_Value = True Then
'''        .Parameters("pPeriodoAdc") = Choose(gsIdioma, "Acumulado - ", "Accrued - ") & Format(CDate(gfMesAct(gsMesAct) & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
'''       Else
'''        .Parameters("pPeriodoAdc") = Format(CDate(gfMesAct(gsMesAct) & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
'''       End If
'''      ']
'''
'''       If Index = 0 Then
'''        .PreviewReport
'''       Else
'''      '[ARREGLAR: Revisar el uso de los tres primeros parámetros de Print.
'''          .Print 1, 0, 0, unCopias
'''      ']ARREGLAR.
'''       End If
'''      .UnLoadReport
'''    End With
'''    Set MRViewer = Nothing
'''  End If
'fin eliminar reporte ***********************************************

    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True


    Set oWBook = oExcel.Workbooks.Add
    Set oSheet = oWBook.Worksheets(1)
    With oSheet
        oSheet.Select
    
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        .Cells(nRowI, 1).Value = "Balance de Comprobación"
        nRowI = nRowI + 2
        Dim x1 As Integer
        Dim x As Integer
        x = 0
        x = x + 1: .Cells(nRowI, x).Value = "Empresa"
        x = x + 1: .Cells(nRowI, x).Value = "Detalle"
        x = x + 1: .Cells(nRowI, x).Value = "Periodo"
        x = x + 1: .Cells(nRowI, x).Value = "Cuenta"
        x = x + 1: .Cells(nRowI, x).Value = "Detalle"
        x = x + 1: .Cells(nRowI, x).Value = "May Debe"
        x = x + 1: .Cells(nRowI, x).Value = "May Haber"
        x = x + 1: .Cells(nRowI, x).Value = "Sdo Deudor"
        x = x + 1: .Cells(nRowI, x).Value = "Sdo Acree"
        
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
    End With
    oExcel.Quit
    Set oExcel = Nothing
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore


  ' elimino el archivo temporal
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
  'ppHabilitacion True


'ini 2015-11-19 rpt consoli x empresa
   porstTmp.Close
   'pocnnTmp.Close
   Set porstTmp = Nothing
   'Set pocnnTmp = Nothing

   pocnnMain.Close
   Set pocnnMain = Nothing
'fin 2015-11-19 rpt consoli x empresa

'ini 2015-12-01 correcc errore
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
    Unload oProgress          ' Unload progress bar window
'fin 2015-12-01 correcc errore

  Exit Sub
Err:
    MsgBox (TEXT_6001)
  If pocnnMain.State = adStateOpen Then
    porstTmp.Close
    'pocnnTmp.Close
    Set porstTmp = Nothing
    'Set pocnnTmp = Nothing
    
'ini 2015-11-19 rpt consoli x empresa
   pocnnMain.Close
   Set pocnnMain = Nothing
'fin 2015-11-19 rpt consoli x empresa
    
  End If

End Sub

'fin 2015-11-20 rpt consoli x empresa

'ini 2015-12-01 balance del mes x cliente
Private Sub pExportaBalCom2(TpoRpt As Integer)

    Dim oProgress As New frmzProgressBar
    oProgress.Show
    oProgress.pgbProgreso.Value = 0: oProgress.pgbProgreso.Min = 0
    oProgress.pgbProgreso.Max = 5
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Min
    oProgress.Caption = "Procesando Ventas"

    Dim pocnnMain As ADODB.Connection
    Set pocnnMain = New ADODB.Connection
    With pocnnMain
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With


    Dim pnNivCta As Byte
        pnNivCta = 9


    Dim chkRango_Value As Boolean
    chkRango_Value = False 'para que no entre en valor
  
    'rangos de mes
    Dim cmbPeriodo_2_ListIndex As Integer
    Dim cmbPeriodo_3_ListIndex As Integer
    cmbPeriodo_2_ListIndex = 0
    cmbPeriodo_3_ListIndex = 12
    
    Dim cmbPeriodo_0 As String
    Dim cmbPeriodo_1 As String
    cmbPeriodo_0 = "2000"
    cmbPeriodo_1 = "2999"
  
  'moneda
  Dim cboTpoMon_ListIndex As Integer
  cboTpoMon_ListIndex = 0 'solo moneda nacional
  
    Dim optAlcance_0_Value As Boolean
    'TpoRpt=1 al mes
    'TpoRpt=2 del mes
    
'ini 2015-12-01 balance del mes x cliente
    'optAlcance_0_Value = True 'para que no entre en valor
    If TpoRpt = 1 Then
    optAlcance_0_Value = True 'para que no entre en valor
    Else
    optAlcance_0_Value = False 'para que no entre en valor
    End If
'fin 2015-12-01 balance del mes x cliente

    
    'txtDato0_Text
    Dim txtDato0_Text As String
    Dim txtDato1_Text As String
    txtDato0_Text = "00"
    txtDato1_Text = "FF"

    Dim chkDivisoria_Value As Boolean
    chkDivisoria_Value = False
    
'********************************************************
  Dim dnContador As Integer, n_Index As Integer
  Dim s_Sentencia As String, s_Sql As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_Moneda As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_MesIni As Integer, n_MesFin As Integer
  Dim l_CreateTB As Boolean
    
  s_AnoIni = Right(IIf(chkRango_Value = vbChecked, cmbPeriodo_0, gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango_Value = vbChecked, cmbPeriodo_1, gsAnoAct), 4)
  
  

  ' Valido el rango de periodos
  If chkRango_Value = vbChecked Then
    s_Mes = Format(cmbPeriodo_2_ListIndex, "00")
    s_Ano = Format(cmbPeriodo_3_ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: 'cmbPeriodo_1.SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: 'cmbPeriodo(3).SetFocus: Exit Sub
  End If
    
  '****ppHabilitacion False
  
  If pnNivCta = 9 Then pnNivCta = Val(Right(gsNivCta, 1))
  s_Moneda = IIf(cboTpoMon_ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
   
    On Error GoTo Err
 
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
  
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore
Dim xxMes As Integer
'xxMes = Val(IIf(TpoRpt = 2, 0, gsMesAct))
For xxMes = Val(IIf(TpoRpt = 2, 0, gsMesAct)) To Val(gsMesAct)
 
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
'ini 2015-12-02 bal compro al mes
'    n_MesIni = Val(IIf(optAlcance_0_Value, 0, gsMesAct))
'    n_MesFin = Val(gsMesAct)
    If TpoRpt = 1 Then
        n_MesIni = Val(IIf(optAlcance_0_Value, 0, gsMesAct))
        n_MesFin = Val(gsMesAct)
    Else
        n_MesIni = Val(IIf(optAlcance_0_Value, 0, xxMes))
        n_MesFin = Val(xxMes)
    End If
'fin 2015-12-02 bal compro al mes
    If chkRango_Value = vbChecked Then
      n_MesIni = Val(IIf(s_Ano = s_AnoIni, cmbPeriodo_2_ListIndex, 1))
      n_MesFin = Val(IIf(s_Ano = s_AnoFin, cmbPeriodo_3_ListIndex, 12))
    End If
    ' Acumulación de saldos
    s_SaldoDeb = "ROUND(": s_SaldoHab = "ROUND("
    For n_Index = n_MesIni To n_MesFin
      s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
    Next n_Index
    s_SaldoDeb = s_SaldoDeb & ", 2)"
    s_SaldoHab = s_SaldoHab & ", 2)"
      
    ' Registros iniciales de saldos
    '2015-11-23 s_Sentencia = "SELECT a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, b.TpoSdo, b.TpoCta, "
    s_Sentencia = "SELECT  a.codemp,CONCAT(a.pdoano,'" & Format(Trim(n_MesFin), "00") & "') pdoano,"
    s_Sentencia = s_Sentencia & " a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, b.TpoSdo, b.TpoCta, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & " AS cSumaD, " & s_SaldoHab & " AS cSumaH, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoDeb & " ELSE 0 END) AS cSumaDt, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoHab & " ELSE 0 END) AS cSumaHt "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCpb ", "")
    '2016-06-02 filtra reportes por EstAct EstAct en Empresa s_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
    s_Sentencia = s_Sentencia & "FROM ((CoCtaAcu a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
'ini 2016-06-02 filtra reportes por EstAct
    's_Sentencia = s_Sentencia & "LEFT JOIN siscfg.tgemp e ON a.codemp=e.codemp) "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & gsNomBDC & ".tgemp e ON a.codemp=e.codemp) "
'fin 2016-06-02 filtra reportes por EstAct
    '2015-11-23 s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "WHERE  "
    '2015-11-23 s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato0_Text & "' AND '" & txtDato1_Text & "' "
    If pnNivCta = 2 Then
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    Else
      If chkDivisoria_Value = 1 Then
        s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=2)) "
      Else
        s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
        s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & " AND b.TpoCta=" & TPOCTA_TRA & ")) "
      End If
    End If
'ini 2016-06-02 filtra reportes por EstAct
      s_Sentencia = s_Sentencia & " AND e.estemp='" & ESTEMPR_ACT & "' "
'fin 2016-06-02 filtra reportes por EstAct
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "HAVING (cSumaD + cSumaH + cSumaDt + cSumaHt) > 0 "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY a.CodCta"
    ' Executo la sentencia
    If Not l_CreateTB Then
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trpRngBceCpb ", "")
      's_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS trpRngBceCpb ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trpRngBceCpb "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
Next xxMes
   
        oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore
   
    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
  With porstTmp
'************************
    .ActiveConnection = pocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    'ini 2015-11-23
    's_Sentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
    s_Sentencia = "SELECT a.codemp,max(e.razemp) razemp,pdoano, "
    s_Sentencia = s_Sentencia & "CodCta, DetCta,  "
    'fin 2015-11-23
    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
    '2015-11-23 s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt "
    s_Sentencia = s_Sentencia & "CASE  WHEN ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2) >0 THEN ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2) ELSE 0 END cSumaDt,"
    s_Sentencia = s_Sentencia & "CASE  WHEN ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2) <0 THEN (ROUND(SUM(cSumaDt), 2)- ROUND(SUM(cSumaHt), 2)) * -1 ELSE 0 END cSumaHt "
    '2015-11-23  s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb a "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & gsNomBDC & ".tgemp e ON  a.codemp=e.codemp "
    '2015-11-23 s_Sentencia = s_Sentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
    s_Sentencia = s_Sentencia & "GROUP BY codemp,pdoano,CodCta, DetCta, TpoSdo, TpoCta "
    s_Sentencia = s_Sentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)) > 0 "
    '2015-11-23 s_Sentencia = s_Sentencia & "ORDER BY CodCta"
    s_Sentencia = s_Sentencia & "ORDER BY codemp,pdoano,CodCta"
    .Source = s_Sentencia
    .Open

'***********************
  End With
  
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore


    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True


    Set oWBook = oExcel.Workbooks.Add
    Set oSheet = oWBook.Worksheets(1)
    With oSheet
        oSheet.Select
    
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        '2015-12-01 balance del mes x cliente .Cells(nRowI, 1).Value = "Balance de Comprobación"
        If TpoRpt = 1 Then
            .Cells(nRowI, 1).Value = "Balance de Comprobación al mes"
        Else
            .Cells(nRowI, 1).Value = "Balance de Comprobación del mes"
        End If
        nRowI = nRowI + 2
        Dim x1 As Integer
        Dim x As Integer
        x = 0
        x = x + 1: .Cells(nRowI, x).Value = "Empresa"
        x = x + 1: .Cells(nRowI, x).Value = "Detalle"
        x = x + 1: .Cells(nRowI, x).Value = "Periodo"
        x = x + 1: .Cells(nRowI, x).Value = "Cuenta"
        x = x + 1: .Cells(nRowI, x).Value = "Detalle"
        x = x + 1: .Cells(nRowI, x).Value = "May Debe"
        x = x + 1: .Cells(nRowI, x).Value = "May Haber"
        x = x + 1: .Cells(nRowI, x).Value = "Sdo Deudor"
        x = x + 1: .Cells(nRowI, x).Value = "Sdo Acree"
        
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
    End With
    oExcel.Quit
    Set oExcel = Nothing
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1 '2015-12-01 correcc errore


    ' elimino el archivo temporal
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")


    porstTmp.Close
    Set porstTmp = Nothing
    
    pocnnMain.Close
    Set pocnnMain = Nothing
    
     oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
     Unload oProgress          ' Unload progress bar window

  Exit Sub
Err:
    MsgBox (TEXT_6001)
    If pocnnMain.State = adStateOpen Then
        porstTmp.Close
        Set porstTmp = Nothing
        pocnnMain.Close
        Set pocnnMain = Nothing
    End If

End Sub
'fin 2015-12-01 balance del mes x cliente

