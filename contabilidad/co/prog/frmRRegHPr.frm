VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRRegHPr 
   Caption         =   "[título]"
   ClientHeight    =   3225
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDiario 
      Caption         =   "Totaliza Diario"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3180
      TabIndex        =   7
      Top             =   855
      Width           =   1335
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Diario"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   1005
      Width           =   4530
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   780
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "frmRRegHPr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
      End
      Begin VB.Label lblDatoDeta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   255
         Width           =   3240
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   19
      Top             =   1800
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   21
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Orden"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   1800
      Width           =   4995
      Begin VB.OptionButton OptOrden 
         Caption         =   "Comprobante"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   18
         Top             =   315
         Width           =   1260
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "Documento"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   2250
         TabIndex        =   17
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "Proveedor"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "Fecha"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   4
      Top             =   135
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6885
         Picture         =   "frmRRegHPr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   1260
      End
      Begin VB.Label lblDatoDeta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   6
         Top             =   255
         Width           =   5520
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   6060
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1035
      Width           =   1260
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7320
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2610
      Width           =   7320
      Begin VB.CommandButton cmdConfig 
         Caption         =   "&Configuración de Impresora"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2355
         TabIndex        =   2
         Top             =   0
         Width           =   1125
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
         Height          =   570
         Left            =   4800
         Picture         =   "frmRRegHPr.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Vista Preliminar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   0
         Left            =   0
         Picture         =   "frmRRegHPr.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1125
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
         Height          =   570
         Index           =   1
         Left            =   1245
         Picture         =   "frmRRegHPr.frx":09D0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
      Begin MSComctlLib.Toolbar toolbar 
         Height          =   600
         Left            =   3600
         TabIndex        =   25
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1058
         ButtonWidth     =   1323
         ButtonHeight    =   1005
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
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A1"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A2"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A3"
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
                  Picture         =   "frmRRegHPr.frx":0AD2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegHPr.frx":0C2C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegHPr.frx":0D86
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegHPr.frx":1148
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegHPr.frx":1812
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   5250
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmRRegHPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MRViewer As MRViewerObject
Attribute MRViewer.VB_VarHelpID = -1

Public udFecha As Date
Public unCopias As Integer
Public unMargenIzquierdo As Integer
Public usDEstino As String
Public usOrientacionRpt As String
Public usOrientacionOri As String
Private paOpciones As Variant
Private pocnnMain As ADODB.Connection
Private porstMRp As ADODB.Recordset

'[Propio del formulario.
Private porstTGAux As ADODB.Recordset
Private porstCodro As ADODB.Recordset
']

Private Sub chkDiario_Click()
  fraRangos.Enabled = (chkDiario.Value = vbUnchecked)
  txtDato(1).Text = IIf(chkDiario.Value = vbChecked, "", txtDato(1).Text)
  lblDatoDeta(1).Caption = IIf(chkDiario.Value = vbChecked, "", lblDatoDeta(1).Caption)
End Sub

Private Sub Form_Load()
   On Error GoTo Err
'ini 2015-07-10 excel hono prof
toolbar.Buttons(1).ButtonMenus(1).Text = "Del Mes"
toolbar.Buttons(1).ButtonMenus(2).Text = "Al Mes"
'fin 2015-07-10 excel hono prof
toolbar.Buttons(1).ButtonMenus(3).Text = "Historico" '2015-09-03 opc historico
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstTGAux = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
   
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstMRp
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TGAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   
   With porstCodro
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
    .Source = .Source & "FROM CODro "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodDro"
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
   End With
 
 ']

 '[Parámetros.                         'Cambiar.
  txtDato.Item(0).DataField = "CodAux"
  txtDato.Item(0).MaxLength = porstTGAux.Fields(txtDato.Item(0).DataField).DefinedSize
 
  txtDato.Item(1).DataField = "CodDro"
  txtDato.Item(1).MaxLength = porstCodro.Fields(txtDato.Item(1).DataField).DefinedSize
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  fraAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  chkDiario.Caption = Choose(gsIdioma, "Totaliza Diario", "Journal Totalizes")
  fraRangos.Caption = Choose(gsIdioma, "Diario", "Journal")
  fraOrden.Caption = Choose(gsIdioma, "Orden", "Order")
  OptOrden(0).Caption = Choose(gsIdioma, "Fecha", "Date")
  OptOrden(1).Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  OptOrden(2).Caption = Choose(gsIdioma, "Documento", "Document")
  OptOrden(3).Caption = Choose(gsIdioma, "Comprobante", "Voucher")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND

 '[Datos predeterminados.              'Cambiar.
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Características de impresión.
   chkImpFecha.Value = vbChecked
   udFecha = Date                      'Fecha en el encabezado.
   unCopias = 1 'frmMain.rptMain.CopiesToPrinter  'Cantidad de Copias.
   unMargenIzquierdo = 240             'Margen izquierdo.
   usDEstino = PRN_DEST_MATR           'PRN_DEST_GRAF:ica _
                                        PRN_DEST_MATR:icial.
   usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                        PRN_ORIE_HORI:zontal.
 ']
   frmOPrnCfg.OrientacionPrn 0, Me
   frmOPrnCfg.lblOriPrn.Caption = Printer.Orientation
   
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub Form_Activate()
   'Orden: Vista Previa, Imprimir, Exportar.
   zaOpciones = Array(gbPms04, gbPms05, gbPms06)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Cambiar. Añadir recordsets.
   porstTGAux.Close
   pocnnMain.Close
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sMoneda As String
     
  ppHabilitacion False
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.FeEDoc, a.RefDoc, b.RucAux, b.RazAux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodDro,'-',a.NroCpb)", "(a.CodDro+'-'+a.NroCpb)") & " AS cDroCpb, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc,'-',a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cNroDoc, "
    .Source = .Source & "a.ImpBru_" & sMoneda & " AS clmBru, a.ImpIR4_" & sMoneda & " AS clmIR4, "
    .Source = .Source & "a.ImpIES_" & sMoneda & " AS clmIES, a.ImpORT_" & sMoneda & " AS clmORT, "
    .Source = .Source & "a.ImpNet_" & sMoneda & " AS clmNet, "
    .Source = .Source & "a.ImpNet_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_TXT, TPOMON_NAC_TXT) & " AS clmpNet_OM, "
    ' Genero la agrupacion del detalle
    If chkDiario.Value = vbChecked Then
    End If
    .Source = .Source & "a.CodDro, c.DetDro, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "a.CodDro", IIf(Trim(txtDato(1).Text) <> "", "a.CodDro", "'drxx'")) & " AS grupo, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "'1'", IIf(Trim(txtDato(1).Text) <> "", "'2'", "'0'")) & " AS resumen "
    .Source = .Source & "FROM ((COHPrDoc a "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    .Source = .Source & "LEFT JOIN CODro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodDro=c.CodDro) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs='" & gsMesAct & "' "
    If Trim(txtDato(0).Text) <> "" Then
      .Source = .Source & "AND a.CodAux = '" & Trim(txtDato(0).Text) & "' "
    End If
    If Trim(txtDato(1).Text) <> "" Then
      .Source = .Source & "AND Left(a.CodDro, " & Len(Trim(txtDato(1).Text)) & ")='" & Trim(txtDato(1).Text) & "' "
    End If
    If OptOrden(0).Value = True Then
      .Source = .Source & "ORDER BY grupo, a.FeEDoc, a.SerDoc, a.NroDoc"
    Else
      If OptOrden(1).Value = True Then
        .Source = .Source & "ORDER BY grupo, a.CodAux, a.FeEDoc, a.SerDoc, a.NroDoc"
      Else
        If OptOrden(2).Value = True Then
          .Source = .Source & "ORDER BY grupo, a.SerDoc, a.NroDoc"
        Else
          .Source = .Source & "ORDER BY grupo, a.CodDro, a.NroCpb"
        End If
      End If
    End If
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRRegHPr.rpt"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      '[ Formula para Simbolo de Moneda ]
      .Formulas(7) = "pSigMon='" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN) & "'"
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRRegHPr.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pSigMon") = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN)
      
      .Parameters("pPagePrinter") = ""
      If porstMRp.RecordCount > 0 Then
        porstMRp.MoveLast
        .Parameters("pPagePrinter") = porstMRp!rucaux & porstMRp!cNroDoc & porstMRp!cDroCpb
        porstMRp.MoveFirst
      End If
      ']
      
      ']
      If Index = 0 Then
        .PreviewReport
      Else
        '[ARREGLAR: Revisar el uso de los tres primeros parámetros de Print.
        .Print 1, 0, 0, unCopias
        ']ARREGLAR.
      End If
      .UnLoadReport
    End With
    Set MRViewer = Nothing
  End If
  ppHabilitacion True

End Sub

Private Sub cmdConfig_Click()
   With frmOPrnCfg
      .ConfiguraPrn 0, Me
   
      .Show vbModal
    
      .ConfiguraPrn 1, Me
   End With
   
   cmdImprimir(1).SetFocus
End Sub

Private Sub cmdSalir_Click()
   frmOPrnCfg.OrientacionPrn 1, Me
   
   Unload Me
End Sub

'ini 2015-07-09 excel hono prof
'Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'
'End Sub

Private Sub toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  'no pinto datos Seleccion.Text = ButtonMenu.Text
  Select Case ButtonMenu.Key
   Case "A1": pExporta 1
   Case "A2": pExporta 2
'   Case "A" & Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
'    pnOpcion = Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
   Case "A3": pExporta 3 '2015-09-03 opc historico
  End Select

End Sub
Private Sub pExporta(TpoRpt As Integer)
'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err

'ini 2015-07-10 excel hono prof
  Dim sMoneda As String
     
  ppHabilitacion False
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)

'fin 2015-07-10 excel hono prof


    Dim pocnnTmp As ADODB.Connection '2014-04-14 Query timeout expired
    Set pocnnTmp = New ADODB.Connection '2014-04-14 Query timeout expired
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
    
    Dim cCadReporte  As String
    Dim sTabla As String
    sTabla = "xlsHPrCab"
    pocnnTmp.Execute fDropTable2(sTabla, 1)

    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
'ini 2015-07-10 excel hono prof
     cCadReporte = cCadReporte & " SELECT "
     cCadReporte = cCadReporte & " concat(a.pdoano,a.mespvs,'00') AS VPERIODO,"
    '.Source = "SELECT a.FeEDoc, a.RefDoc, b.RucAux, b.RazAux, "
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodDro,'-',a.NroCpb)", "(a.CodDro+'-'+a.NroCpb)") & " AS cDroCpb, "
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc,'-',a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cNroDoc, "
    ''2015-12-17 adicion ref  cCadReporte = cCadReporte & "a.FeEDoc, a.RefDoc, b.RucAux, b.RazAux, "
    cCadReporte = cCadReporte & "a.FeEDoc, b.RucAux, b.RazAux, "
    cCadReporte = cCadReporte & "a.ImpBru_" & sMoneda & " AS clmBru, a.ImpIR4_" & sMoneda & " AS clmIR4, "
    cCadReporte = cCadReporte & "a.ImpIES_" & sMoneda & " AS clmIES, a.ImpORT_" & sMoneda & " AS clmORT, "
    cCadReporte = cCadReporte & "a.ImpNet_" & sMoneda & " AS clmNet, "
    cCadReporte = cCadReporte & "a.ImpNet_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_TXT, TPOMON_NAC_TXT) & " AS clmpNet_OM, "
    ' Genero la agrupacion del detalle
    If chkDiario.Value = vbChecked Then
    End If
    cCadReporte = cCadReporte & "a.TpoMon, " '2015-05-14
    cCadReporte = cCadReporte & "GloDoc, " '2015-06-04 adicion glodoc
    cCadReporte = cCadReporte & "ifnull(a.refdoc,'') refdoc, " '2015-12-17 adicion ref
    cCadReporte = cCadReporte & "a.CodDro, c.DetDro, "
    cCadReporte = cCadReporte & IIf(chkDiario.Value = vbChecked, "a.CodDro", IIf(Trim(txtDato(1).Text) <> "", "a.CodDro", "'drxx'")) & " AS grupo, "
    cCadReporte = cCadReporte & IIf(chkDiario.Value = vbChecked, "'1'", IIf(Trim(txtDato(1).Text) <> "", "'2'", "'0'")) & " AS resumen "
    cCadReporte = cCadReporte & "FROM ((COHPrDoc a "
    cCadReporte = cCadReporte & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    cCadReporte = cCadReporte & "LEFT JOIN CODro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodDro=c.CodDro) "
    cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
'    cCadReporte = cCadReporte & "AND a.pdoano='" & gsAnoAct & "' "
'    cCadReporte = cCadReporte & "AND a.MesPvs='" & gsMesAct & "' "
    If TpoRpt = 1 Then
        cCadReporte = cCadReporte & "AND a.pdoano='" & gsAnoAct & "' "
        cCadReporte = cCadReporte & "AND a.MesPvs='" & gsMesAct & "' "
    'Else '2015-09-03 opc historico
'ini 2015-09-03 opc historico
    ElseIf TpoRpt = 2 Then
        cCadReporte = cCadReporte & "AND  concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
        cCadReporte = cCadReporte & "  concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "'  "
    Else
        'cCadReporte = cCadReporte & "AND  concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
        cCadReporte = cCadReporte & "AND  concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "'  "
    End If
'fin 2015-09-03 opc historico
'2015-09-03 ya existe este pedazo de codigo aqui
    If Trim(txtDato(0).Text) <> "" Then
      cCadReporte = cCadReporte & "AND a.CodAux = '" & Trim(txtDato(0).Text) & "' "
    End If
    If Trim(txtDato(1).Text) <> "" Then
      cCadReporte = cCadReporte & "AND Left(a.CodDro, " & Len(Trim(txtDato(1).Text)) & ")='" & Trim(txtDato(1).Text) & "' "
    End If
'ini 2015-01-09 adiciona ruc
'2015-09-03 ya existe este pedazo de codigo
''      If Trim(txtDato(0).Text) <> "" Then
''          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato(0).Text) & "' "
''      End If
'fin 2015-01-09 adiciona ruc
    
    If OptOrden(0).Value = True Then
      cCadReporte = cCadReporte & "ORDER BY grupo, a.FeEDoc, a.SerDoc, a.NroDoc"
    Else
      If OptOrden(1).Value = True Then
        cCadReporte = cCadReporte & "ORDER BY grupo, a.CodAux, a.FeEDoc, a.SerDoc, a.NroDoc"
      Else
        If OptOrden(2).Value = True Then
          cCadReporte = cCadReporte & "ORDER BY grupo, a.SerDoc, a.NroDoc"
        Else
          cCadReporte = cCadReporte & "ORDER BY grupo, a.CodDro, a.NroCpb"
        End If
      End If
    End If
    pocnnTmp.Execute cCadReporte

''    cCadReporte = cCadReporte & "SELECT"
''    cCadReporte = cCadReporte & "    concat(a.pdoano,a.mespvs,'00') AS VPERIODO,"
''    cCadReporte = cCadReporte & "    concat(a.CodDro,a.NroCpb) as VNUMREGOPE,"
''    cCadReporte = cCadReporte & "    date_format(a.feedoc,'%d/%m/%Y')as VFECCOM,"
''    cCadReporte = cCadReporte & "    date_format(a.FevDOC,'%d/%m/%Y')as VFECVENPAG,"
''    cCadReporte = cCadReporte & "    b.CodTDc as VTIPDOCCOM, a.SerDoc AS VNUMSER, a. NroDoc AS VNUMDOCCOI,"
''    cCadReporte = cCadReporte & "    IF(ifnull(a.NroDoc_Fin,''),a.NroDoc_Fin,'0') AS VNUMDOCCOF,"
''    cCadReporte = cCadReporte & "    MID(c.tpodci,2,1) AS VTIPDIDCLI,"
''    cCadReporte = cCadReporte & "    c.Codaux AS VNUMDIDCLI,"
''    cCadReporte = cCadReporte & "    replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
''    cCadReporte = cCadReporte & "    ifnull(MID(c.RazAux,1,60) ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
''    cCadReporte = cCadReporte & "    as VAPENOMRSO,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpExp_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VVALFACEXP,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VBASIMPGRA,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTEXO,"
''    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTINA,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VISC,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIGVIPM,"
''    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VBASIMIVAP,"
''    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIVAP,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VOTRTRICGO,"
''    cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTCOM,"
''    cCadReporte = cCadReporte & "    format(a.imptcb,3) * 1 AS VTIPCAM,"
''    cCadReporte = cCadReporte & "    IF(ifnull(codtdc_ref,''),date_format(feedoc_ref,'%d/%m/%Y'),'01/01/0001') as VFECCOMMOD,"
''    cCadReporte = cCadReporte & "    IF(ifnull(a.codtdc_ref,''),a.codtdc_ref,'00') as VTIPCCOMOD,"
''    cCadReporte = cCadReporte & "    IF(ifnull(a.serdoc_ref,''),a.serdoc_ref,'-')  as VNUMSERMOD,"
''    cCadReporte = cCadReporte & "    IF(ifnull(a.nrodoc_ref,''),a.nrodoc_ref,'-')  as VNUMCOMMOD,"
''    cCadReporte = cCadReporte & "    IF(a.ImpTot_MN <>0.00,'1','2') as VESTOPE,"
''    cCadReporte = cCadReporte & "    '' AS VINTDIAMAY,"
''    cCadReporte = cCadReporte & "    '' AS VINTKARDEX,"
''    cCadReporte = cCadReporte & "    '' AS VINTREG "
''    cCadReporte = cCadReporte & "    ,a.TpoMon " '2015-05-14
''    cCadReporte = cCadReporte & "    ,GloDoc " '2015-06-04 adicion glodoc
''
''    cCadReporte = cCadReporte & "FROM (((COVtaDoc a "
''    cCadReporte = cCadReporte & "LEFT JOIN TGTDc b ON  a.codemp=b.codemp and a.CodTDc=b.CodTDc) "
''    cCadReporte = cCadReporte & "LEFT JOIN TGAux c ON  a.codemp=c.codemp  and a.CodAux=c.CodAux) "
''    cCadReporte = cCadReporte & "LEFT JOIN CODro d ON  a.codemp=d.codemp  and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
''    'cCadReporte = cCadReporte & "WHERE a.codemp='001' and a.pdoano='2012' and a.Mespvs >='01'  and a.Mespvs <='05' AND IFNULL(a.CodAux, '')<>'' AND IFNULL(a.CodDro, '')<>'' "
''    cCadReporte = cCadReporte & "WHERE "
''    cCadReporte = cCadReporte & "   a.codemp='" & gsCodEmp & "' and "
'''    cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & gsAnoAct & Left(cmbEjercicio.Text, 2) & "' AND "
''        If TpoRpt = 1 Then
''            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' AND "
''        Else
''            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
''            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND  "
''        End If
''    cCadReporte = cCadReporte & "   IFNULL(a.CodAux, '')<>'' AND  "
''    cCadReporte = cCadReporte & "   IFNULL(a.CodDro, '')<>'' "
''    cCadReporte = cCadReporte & "ORDER BY a.mespvs ,a.CodTDc, a.SerDoc, a.NroDoc  ASC "

'fin 2015-07-10 excel hono prof

    
    
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
        
        .Cells(nRowI, 1).Value = "Registro de Honorarios Profesionales"
        nRowI = nRowI + 2
        Dim x1 As Integer
        .Cells(nRowI, 1).Value = "Periodo"
        .Cells(nRowI, 2).Value = "NºComprob."
        .Cells(nRowI, 3).Value = "Documento"
        .Cells(nRowI, 4).Value = "Fecha"
        '2015-12-17 adicion ref .Cells(nRowI, 5).Value = "Referencia"
        .Cells(nRowI, 6 - 1).Value = "RUC"
        .Cells(nRowI, 7 - 1).Value = "R.Social"
        .Cells(nRowI, 8 - 1).Value = "Bas. Imp."
        .Cells(nRowI, 9 - 1).Value = "I.R.4ta C."
        .Cells(nRowI, 10 - 1).Value = "I.E.S"
        .Cells(nRowI, 11 - 1).Value = "Otros"
        .Cells(nRowI, 12 - 1).Value = "Total"
        .Cells(nRowI, 13 - 1).Value = "Total $"
        .Cells(nRowI, 14 - 1).Value = "TpoMon"
        .Cells(nRowI, 15 - 1).Value = "Glosa"
        .Cells(nRowI, 15).Value = "Referencia" '2015-12-17 adicion ref
        
        .Cells(nRowI, 16).Value = "Diario"
        .Cells(nRowI, 17).Value = "Detalle"
        .Cells(nRowI, 18).Value = "Grupo"
        .Cells(nRowI, 19).Value = "Resumen"
        
'        .Cells(nRowI, 1).Value = "Periodo"
'        .Cells(nRowI, 2).Value = "Nº Reg."
'        .Cells(nRowI, 3).Value = "F.Vta"
'        .Cells(nRowI, 4).Value = "F. Pago"
'        .Cells(nRowI, 5).Value = "T.Doc"
'        .Cells(nRowI, 6).Value = "Serie"
'        .Cells(nRowI, 7).Value = "VNUMDOCCCOI"
'        .Cells(nRowI, 8).Value = "Nº Doc."
'        .Cells(nRowI, 9).Value = "Tpo.Cli"
'        .Cells(nRowI, 10).Value = "RUC"
'        .Cells(nRowI, 11).Value = "R.Social"
'        .Cells(nRowI, 12).Value = "VVALFACEXP"
'        .Cells(nRowI, 13).Value = "VBASIMPGRA"
'        .Cells(nRowI, 14).Value = "VIMPTOTEXO"
'        .Cells(nRowI, 15).Value = "VIMPTOTINA"
'        .Cells(nRowI, 16).Value = "VISC"
'        .Cells(nRowI, 17).Value = "VIGVIPM"
'        .Cells(nRowI, 18).Value = "VBASIMIVAP"
'        .Cells(nRowI, 19).Value = "VIVAP"
'        .Cells(nRowI, 20).Value = "VOTRTRICGO"
'        .Cells(nRowI, 21).Value = "CIMPTOTCOM"
'        .Cells(nRowI, 22).Value = "VTIPCAM"
'        .Cells(nRowI, 23).Value = "VFECCOMMOD"
'        .Cells(nRowI, 24).Value = "VTIPCCOMOD"
'        .Cells(nRowI, 25).Value = "VNUMSERMOD"
'        .Cells(nRowI, 26).Value = "VNUMCOMMOD"
'        .Cells(nRowI, 27).Value = "VESTOPE"
'        .Cells(nRowI, 28).Value = "VINTDIAMAY"
'        .Cells(nRowI, 29).Value = "VINTKARDEX"
'        .Cells(nRowI, 30).Value = "VINTREG"
'        .Cells(nRowI, 31).Value = "TpoMon"
'        .Cells(nRowI, 32).Value = "Glosa" '2015-06-04 adicion glodoc

        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
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

  ppHabilitacion True

'fin exporta datos a excel

   porstTmp.Close
   pocnnTmp.Close
   Set porstTmp = Nothing
   Set pocnnTmp = Nothing

  Exit Sub
Err:
    MsgBox (TEXT_6001)
  If pocnnTmp.State = adStateOpen Then
    porstTmp.Close
    pocnnTmp.Close
    Set porstTmp = Nothing
    Set pocnnTmp = Nothing
  End If

End Sub

'fin 2015-07-09 excel hono prof

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)

  Select Case Index    'Busca el dato en su tabla principal.
   Case 0                              'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
   Case 1
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    If Len(txtDato(Index)) <> 4 Then txtDato(Index).SetFocus: Exit Sub
  End Select

End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
    Case 0                              'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    Case 1
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   
   Select Case tnIndex                 'Cambiar.
    Case 0
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstTGAux
         .MoveFirst
         .Find "CodAux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !razAux
         End If
      End With
    Case 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstCodro
         .MoveFirst
         .Find "Coddro='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
         End If
      End With
   End Select

End Function

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

  'Controles del formulario.
'   cboTpoMon.Enabled = tbHabilitar
'   dtpFecha.Enabled = tbHabilitar
'   optTipo(0).Enabled = tbHabilitar
'   optTipo(1).Enabled = tbHabilitar
'   With txtDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With cmdDatoAyud
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With lblDatoDeta
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

