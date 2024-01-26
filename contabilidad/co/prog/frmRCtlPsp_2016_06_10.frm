VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRCtlPsp 
   Caption         =   "[título]"
   ClientHeight    =   4380
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   305
      Index           =   0
      Left            =   6720
      Picture         =   "frmRCtlPsp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtLlave 
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
      TabIndex        =   22
      Top             =   360
      Width           =   950
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   18
      Top             =   2880
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   19
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   16
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   6990
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRCtlPsp.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   915
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRCtlPsp.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   555
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
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   900
         Width           =   945
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
         Left            =   135
         TabIndex        =   4
         Top             =   540
         Width           =   945
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
         Left            =   1065
         TabIndex        =   11
         Top             =   900
         Width           =   5520
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
         Left            =   1065
         TabIndex        =   10
         Top             =   555
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.PictureBox picOpciones 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   6990
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3720
      Width           =   6990
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
         Picture         =   "frmRCtlPsp.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmRCtlPsp.frx":0600
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmRCtlPsp.frx":0B32
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
      Begin MSComctlLib.Toolbar toolbar 
         Height          =   570
         Left            =   3600
         TabIndex        =   26
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1005
         ButtonWidth     =   1429
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Otros Rpt"
               Object.ToolTipText     =   "Otros Reportes de presuesto"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A1"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A2"
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
                  Picture         =   "frmRCtlPsp.frx":0C7C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRCtlPsp.frx":0DD6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRCtlPsp.frx":0F30
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRCtlPsp.frx":12F2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRCtlPsp.frx":19BC
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label lblLlaveDeta 
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
      Left            =   1080
      TabIndex        =   25
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "C.Costos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   17
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   705
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   4800
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
End
Attribute VB_Name = "frmRCtlPsp"
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
Private porstMRpRs      As Recordset
Private porstCOCta  As ADODB.Recordset
Private pnNivCta    As Byte
']

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      'txtLlave(Index).SetFocus
   End Select
   ppAyuBusx Index
End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
'ini sql8 2015-03-23
toolbar.Buttons(1).ButtonMenus(1).Text = "Cuenta x Presupuesto Vista Previa"
toolbar.Buttons(1).ButtonMenus(2).Text = "Cuenta x Presupuesto Impresion directa"
'fin sql8 2015-03-23
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
    Set pocnnMain = New ADODB.Connection
    Set porstMRp = New ADODB.Recordset
    Set porstCOCta = New ADODB.Recordset
    Set porstMRpRs = New ADODB.Recordset
   
    With pocnnMain
        .CursorLocation = adUseClient
        .ConnectionString = CONNSTRG & gsNomBDS
        .Open
    End With
    With porstMRp
        .ActiveConnection = pocnnMain
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
    End With
    With porstCOCta
        .ActiveConnection = pocnnMain
        .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
        .Source = .Source & "FROM CoCta "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
    End With
    With porstMRpRs
        .ActiveConnection = pocnnMain
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Source = "SELECT "
        .Source = .Source & " codcta,detcta,ordrep_1,detordrep_1,ordrep_2,detordrep_2,impsdo_mes,impsdo_acu,pspmes,"
        .Source = .Source & " pspacu,numcol0,numcol1,numcol2,numcol3,numcol4,numcol5,numcol6,numcol7,numcol8,numcol9,numcol10,"
        .Source = .Source & " numcol11,numcol12,codaux,razaux,rucaux,codtdc,abvtdc,serdoc,nrodoc,feedoc,fevdoc,refdoc,gloite,codcco,"
        .Source = .Source & " cdocum , coddro, nrocpb, mespvs, blqite, fehope, nomrpt, cdrocpb, tpocta, tposdo, usrcre, detcco,codemp,pdoano"
        .Source = .Source & " FROM COTmpRpt "
        .Source = .Source & " WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & " AND UsrCre='" & gsCodUsr & "' "
        .Source = .Source & " AND NomRpt='rptRCtlPsp' "
        .Source = .Source & " ORDER BY OrdRep_1, OrdRep_2, CodCta,codcco"
'        .Open
    End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Currency :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
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
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !CodCta
      .MoveFirst
      txtDato(0).Text = !CodCta
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   
  'Otros.
   
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
   porstCOCta.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstMRp = Nothing
   Set porstMRpRs = Nothing
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
    proc_cmd_imprimir Index, 0
End Sub

Private Sub proc_cmd_imprimir(Index As Integer, ptpo_rpt As Integer)
  Dim sCadena, sIni, cCadReporte As String
  Dim nContador As Integer
  
  ppHabilitacion False
         
  sIni = IIf(OptTipo(0).Value, "", "SUM")
  sCadena = sIni & "("
  For nContador = 1 To Val(gsMesAct)
    sCadena = sCadena & "Imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & Format(nContador, "00") & IIf(nContador = Val(gsMesAct), ")", " + ")
  Next nContador
  sCadena = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & sCadena & ", 0), 2)"
    
  If txtLlave(0).Text <> "" Then
    
  With porstMRp
    If .State = adStateOpen Then .Close
    If OptTipo(0).Value Then
      .Source = "SELECT a.CodCta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(a.CodCta,2) AS cTitulo2, "
    Else
      .Source = "SELECT LEFT(a.CodCta, 2) AS cCuenta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
    End If
    .Source = .Source & "a.codcco AS codcco, "
    .Source = .Source & "LEFT(a.OrdRep,1) AS cOrden, "
    .Source = .Source & "RIGHT(a.OrdRep,2) AS cOrden1, "
    .Source = .Source & "(CASE LEFT(a.OrdRep,1) WHEN '" & TPOGRU1_TXT_0 & "' THEN '" & TPOGRU1_TXT_1 & "' WHEN '" & TPOGRU2_TXT_0 & "' THEN '" & TPOGRU2_TXT_1 & "' ELSE '" & TPOGRU2_TXT_1 & "' END) AS cTitulo, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpAcuDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpAcuHaber, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpMesDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpMesHaber, "
    .Source = .Source & "ROUND(" & sIni & "(CASE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " WHEN 0 THEN 0.00 ELSE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " END), 2) AS cPreMes, "
    .Source = .Source & sCadena & " AS cPreAcu "
    .Source = .Source & "FROM ((COPsp a "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta) "
    If OptTipo(0).Value Then
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    Else
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.CodCta,2)=b.CodCta) "
    End If
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "AND a.codcco='" & txtLlave(0).Text & "' "
    If OptTipo(0).Value Then
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & "ORDER BY a.OrdRep, a.CodCta"
    Else
      .Source = .Source & "GROUP BY LEFT(a.CodCta, 2), a.OrdRep,  b.DetCta "
      .Source = .Source & "ORDER BY a.OrdRep, LEFT(a.CodCta, 2)"
    End If
    .Open
  End With
  
  Else
  
  With porstMRp
    If .State = adStateOpen Then .Close
    If OptTipo(0).Value Then
      .Source = "SELECT a.CodCta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(a.CodCta,2) AS cTitulo2, "
    Else
      .Source = "SELECT LEFT(a.CodCta, 2) AS cCuenta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
    End If
    .Source = .Source & "a.codcco AS codcco, "
    .Source = .Source & "LEFT(a.OrdRep,1) AS cOrden, "
    .Source = .Source & "RIGHT(a.OrdRep,2) AS cOrden1, "
    .Source = .Source & "(CASE LEFT(a.OrdRep,1) WHEN '" & TPOGRU1_TXT_0 & "' THEN '" & TPOGRU1_TXT_1 & "' WHEN '" & TPOGRU2_TXT_0 & "' THEN '" & TPOGRU2_TXT_1 & "' ELSE '" & TPOGRU2_TXT_1 & "' END) AS cTitulo, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpAcuDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpAcuHaber, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpMesDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpMesHaber, "
    .Source = .Source & "ROUND(" & sIni & "(CASE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " WHEN 0 THEN 0.00 ELSE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " END), 2) AS cPreMes, "
    .Source = .Source & "" & sCadena & " AS cPreAcu "
    .Source = .Source & "FROM ((COPsp a "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta) "
    If OptTipo(0).Value Then
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    Else
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.CodCta,2)=b.CodCta) "
    End If
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    If OptTipo(0).Value Then
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & " GROUP BY a.OrdRep, a.CodCta"
      .Source = .Source & " ORDER BY a.OrdRep, a.CodCta"
    Else
      .Source = .Source & "GROUP BY LEFT(a.CodCta, 2), a.OrdRep,  b.DetCta "
      .Source = .Source & "ORDER BY a.OrdRep, LEFT(a.CodCta, 2)"
    End If
    .Open
  End With
  
  
  End If
  
  
  If OptTipo(0).Value Then Llena_Temporal
         
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " -" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, IIf(OptTipo(0).Value, porstMRpRs, porstMRp)
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".rpt"
      ' .WindowShowGroupTree = True
      'Fórmular propias.
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    
    With MRViewer
      If OptTipo(0).Value Then
        .DataRecordSet = porstMRpRs
      Else
        .DataRecordSet = porstMRp
      End If
      .LoadReport gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " -" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
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
  If OptTipo(0).Value Then porstMRpRs.Close
  
  ppHabilitacion True
End Sub


Private Sub cmdImprimir_Click_2016_06_09_nvo_rpt(Index As Integer)
  Dim sCadena, sIni, cCadReporte As String
  Dim nContador As Integer
  
  ppHabilitacion False
         
  sIni = IIf(OptTipo(0).Value, "", "SUM")
  sCadena = sIni & "("
  For nContador = 1 To Val(gsMesAct)
    sCadena = sCadena & "Imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & Format(nContador, "00") & IIf(nContador = Val(gsMesAct), ")", " + ")
  Next nContador
  sCadena = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & sCadena & ", 0), 2)"
    
  If txtLlave(0).Text <> "" Then
    
  With porstMRp
    If .State = adStateOpen Then .Close
    If OptTipo(0).Value Then
      .Source = "SELECT a.CodCta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(a.CodCta,2) AS cTitulo2, "
    Else
      .Source = "SELECT LEFT(a.CodCta, 2) AS cCuenta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
    End If
    .Source = .Source & "a.codcco AS codcco, "
    .Source = .Source & "LEFT(a.OrdRep,1) AS cOrden, "
    .Source = .Source & "RIGHT(a.OrdRep,2) AS cOrden1, "
    .Source = .Source & "(CASE LEFT(a.OrdRep,1) WHEN '" & TPOGRU1_TXT_0 & "' THEN '" & TPOGRU1_TXT_1 & "' WHEN '" & TPOGRU2_TXT_0 & "' THEN '" & TPOGRU2_TXT_1 & "' ELSE '" & TPOGRU2_TXT_1 & "' END) AS cTitulo, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpAcuDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpAcuHaber, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpMesDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpMesHaber, "
    .Source = .Source & "ROUND(" & sIni & "(CASE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " WHEN 0 THEN 0.00 ELSE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " END), 2) AS cPreMes, "
    .Source = .Source & sCadena & " AS cPreAcu "
    .Source = .Source & "FROM ((COPsp a "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta) "
    If OptTipo(0).Value Then
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    Else
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.CodCta,2)=b.CodCta) "
    End If
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "AND a.codcco='" & txtLlave(0).Text & "' "
    If OptTipo(0).Value Then
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & "ORDER BY a.OrdRep, a.CodCta"
    Else
      .Source = .Source & "GROUP BY LEFT(a.CodCta, 2), a.OrdRep,  b.DetCta "
      .Source = .Source & "ORDER BY a.OrdRep, LEFT(a.CodCta, 2)"
    End If
    .Open
  End With
  
  Else
  
  With porstMRp
    If .State = adStateOpen Then .Close
    If OptTipo(0).Value Then
      .Source = "SELECT a.CodCta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(a.CodCta,2) AS cTitulo2, "
    Else
      .Source = "SELECT LEFT(a.CodCta, 2) AS cCuenta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
    End If
    .Source = .Source & "a.codcco AS codcco, "
    .Source = .Source & "LEFT(a.OrdRep,1) AS cOrden, "
    .Source = .Source & "RIGHT(a.OrdRep,2) AS cOrden1, "
    .Source = .Source & "(CASE LEFT(a.OrdRep,1) WHEN '" & TPOGRU1_TXT_0 & "' THEN '" & TPOGRU1_TXT_1 & "' WHEN '" & TPOGRU2_TXT_0 & "' THEN '" & TPOGRU2_TXT_1 & "' ELSE '" & TPOGRU2_TXT_1 & "' END) AS cTitulo, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpAcuDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpAcuHaber, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpMesDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpMesHaber, "
    .Source = .Source & "ROUND(" & sIni & "(CASE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " WHEN 0 THEN 0.00 ELSE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " END), 2) AS cPreMes, "
    .Source = .Source & "" & sCadena & " AS cPreAcu "
    .Source = .Source & "FROM ((COPsp a "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta) "
    If OptTipo(0).Value Then
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    Else
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.CodCta,2)=b.CodCta) "
    End If
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    If OptTipo(0).Value Then
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & " GROUP BY a.OrdRep, a.CodCta"
      .Source = .Source & " ORDER BY a.OrdRep, a.CodCta"
    Else
      .Source = .Source & "GROUP BY LEFT(a.CodCta, 2), a.OrdRep,  b.DetCta "
      .Source = .Source & "ORDER BY a.OrdRep, LEFT(a.CodCta, 2)"
    End If
    .Open
  End With
  
  
  End If
  
  
  If OptTipo(0).Value Then Llena_Temporal
         
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " -" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, IIf(OptTipo(0).Value, porstMRpRs, porstMRp)
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".rpt"
      ' .WindowShowGroupTree = True
      'Fórmular propias.
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    
    With MRViewer
      If OptTipo(0).Value Then
        .DataRecordSet = porstMRpRs
      Else
        .DataRecordSet = porstMRp
      End If
      .LoadReport gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " -" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
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
  If OptTipo(0).Value Then porstMRpRs.Close
  
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

'ini 2015-05-04
Private Sub toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  'no pinto datos Seleccion.Text = ButtonMenu.Text
  Select Case ButtonMenu.Key
   Case "A1": pReporte 0
   Case "A2": pReporte 1
'   Case "A" & Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
'    pnOpcion = Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
  End Select

End Sub
'fin 2015-05-04

'ini 2015-05-04
Private Sub pReporte(Index As Integer)
'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err
'++++++++ini adiciona codigo
  Dim sCadena, sIni, cCadReporte As String
  Dim nContador As Integer
  
  ppHabilitacion False
         
  sIni = IIf(OptTipo(0).Value, "", "SUM")
  sCadena = sIni & "("
  For nContador = 1 To Val(gsMesAct)
    sCadena = sCadena & "Imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & Format(nContador, "00") & IIf(nContador = Val(gsMesAct), ")", " + ")
  Next nContador
  sCadena = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & sCadena & ", 0), 2)"
    
'ini 2015-05-04 ctr.ppto x cta
'hemos encontrado que:
'If txtLlave(0).Text <> "" Then
'se difere con las instrucciones del else solo en la parate
'de filtro de centro de costo:
'    .Source = .Source & "AND a.codcco='" & txtLlave(0).Text & "' "
'por este motivo hemos decidido juntar en un sola logica ambos y adicinar condicion para lo del ccosto
'fin 2015-05-04 ctr.ppto x cta
    
'''  If txtLlave(0).Text <> "" Then
'''
'''  With porstMRp
'''    If .State = adStateOpen Then .Close
'''    If OptTipo(0).Value Then
'''      .Source = "SELECT a.CodCta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(a.CodCta,2) AS cTitulo2, "
'''    Else
'''      .Source = "SELECT LEFT(a.CodCta, 2) AS cCuenta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
'''    End If
'''    .Source = .Source & "a.codcco AS codcco, "
'''    .Source = .Source & "LEFT(a.OrdRep,1) AS cOrden, "
'''    .Source = .Source & "RIGHT(a.OrdRep,2) AS cOrden1, "
'''    .Source = .Source & "(CASE LEFT(a.OrdRep,1) WHEN '" & TPOGRU1_TXT_0 & "' THEN '" & TPOGRU1_TXT_1 & "' WHEN '" & TPOGRU2_TXT_0 & "' THEN '" & TPOGRU2_TXT_1 & "' ELSE '" & TPOGRU2_TXT_1 & "' END) AS cTitulo, "
'''    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpAcuDebe, "
'''    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpAcuHaber, "
'''    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpMesDebe, "
'''    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpMesHaber, "
'''    .Source = .Source & "ROUND(" & sIni & "(CASE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " WHEN 0 THEN 0.00 ELSE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " END), 2) AS cPreMes, "
'''    .Source = .Source & sCadena & " AS cPreAcu "
'''    .Source = .Source & "FROM ((COPsp a "
'''    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta) "
'''    If OptTipo(0).Value Then
'''      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
'''    Else
'''      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.CodCta,2)=b.CodCta) "
'''    End If
'''    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
'''    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
'''    .Source = .Source & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
'''    .Source = .Source & "AND a.codcco='" & txtLlave(0).Text & "' "
'''    If OptTipo(0).Value Then
'''      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
'''      .Source = .Source & "ORDER BY a.OrdRep, a.CodCta"
'''    Else
'''      .Source = .Source & "GROUP BY LEFT(a.CodCta, 2), a.OrdRep,  b.DetCta "
'''      .Source = .Source & "ORDER BY a.OrdRep, LEFT(a.CodCta, 2)"
'''    End If
'''    .Open
'''  End With
'''
'''  Else
  
  With porstMRp
    If .State = adStateOpen Then .Close
    If OptTipo(0).Value Then
'ini 2015-05-04 ctr.ppto x cta
      '.Source = "SELECT a.CodCta, a.OrdRep, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(a.CodCta,2) AS cTitulo2, "
      .Source = "SELECT COCtaAcu.CodCta, " & fIsNull() & "a.OrdRep,'0999') OrdRep, "
      .Source = .Source & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, LEFT(COCtaAcu.CodCta,2) AS cTitulo2, "
'fin 2015-05-04 ctr.ppto x cta
    Else
      .Source = "SELECT LEFT(COCtaAcu.CodCta, 2) AS cCuenta," & fIsNull() & "a.OrdRep,'0999') OrdRep, "
      .Source = .Source & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, "
    End If
    .Source = .Source & "a.codcco AS codcco, "
    .Source = .Source & fIsNull() & "LEFT(a.OrdRep,1),'0') AS cOrden, "
    .Source = .Source & fIsNull() & "RIGHT(a.OrdRep,2),'99') AS cOrden1, "
    .Source = .Source & "(CASE LEFT(a.OrdRep,1) WHEN '" & TPOGRU1_TXT_0 & "' THEN '" & TPOGRU1_TXT_1 & "' WHEN '" & TPOGRU2_TXT_0 & "' THEN '" & TPOGRU2_TXT_1 & "' WHEN '" & TPOGRU3_TXT_0 & "' THEN '" & TPOGRU3_TXT_1 & "' ELSE '" & TPOGRU4_TXT_1 & "' END) AS cTitulo, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpAcuDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & " + " & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpAcuHaber, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & "), 2) AS clmpMesDebe, "
    .Source = .Source & "ROUND(" & sIni & "(" & gsAcuMes(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & "), 2) AS clmpMesHaber, "
    .Source = .Source & "ROUND(" & fIsNull() & sIni & "(CASE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " WHEN 0 THEN 0.00 ELSE imp" & IIf(cboTpoMon.ListIndex = 0, "MN_", "ME_") & gsMesAct & " END),0), 2) AS cPreMes, "
    .Source = .Source & "" & sCadena & " AS cPreAcu "
    .Source = .Source & "FROM ((COCtaAcu "
    '2015-05-04 ctr.ppto x cta.Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta) "
    .Source = .Source & "LEFT JOIN COPsp a ON COCtaAcu.codemp=a.codemp AND COCtaAcu.pdoano=a.pdoano AND COCtaAcu.CodCta=a.CodCta)"
    If OptTipo(0).Value Then
      '2015-05-04 ctr.ppto x cta.Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
      .Source = .Source & "LEFT JOIN CoCta b ON COCtaAcu.codemp=b.codemp AND COCtaAcu.pdoano=b.pdoano AND COCtaAcu.CodCta=b.CodCta) "
    Else
      .Source = .Source & "LEFT JOIN CoCta b ON COCtaAcu.codemp=b.codemp AND COCtaAcu.pdoano=b.pdoano AND LEFT(COCtaAcu.CodCta,2)=b.CodCta) "
    End If
    .Source = .Source & "WHERE COCtaAcu.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND COCtaAcu.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND COCtaAcu.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
'ini 2015-05-07 adicion centro de costos
    If txtLlave(0).Text <> "" Then
    .Source = .Source & "AND a.codcco='" & txtLlave(0).Text & "' "
    End If
'fin 2015-05-07 adicion centro de costos
    
    If OptTipo(0).Value Then
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & " GROUP BY a.OrdRep, COCtaAcu.CodCta"
      .Source = .Source & " ORDER BY a.OrdRep, COCtaAcu.CodCta"
    Else
      .Source = .Source & "GROUP BY LEFT(COCtaAcu.CodCta, 2), a.OrdRep,  b.DetCta "
      .Source = .Source & "ORDER BY a.OrdRep, LEFT(COCtaAcu.CodCta, 2)"
    End If
    .Open
  End With
  
  
'''  End If
  
  
  If OptTipo(0).Value Then Llena_Temporal
         
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " -" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, IIf(OptTipo(0).Value, porstMRpRs, porstMRp)
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".rpt"
      ' .WindowShowGroupTree = True
      'Fórmular propias.
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    
    With MRViewer
      If OptTipo(0).Value Then
        .DataRecordSet = porstMRpRs
      Else
        .DataRecordSet = porstMRp
      End If
      .LoadReport gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " -" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
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
  If OptTipo(0).Value Then porstMRpRs.Close
  
  ppHabilitacion True

'++++++++fin adiciona codigo
  Exit Sub
Err:
    MsgBox (TEXT_6001)
'  If pocnnTmp.State = adStateOpen Then
'    porstTmp.Close
'    pocnnTmp.Close
'    Set porstTmp = Nothing
'    Set pocnnTmp = Nothing
'  End If

End Sub
'fin 2015-05-04

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
   'Select Case Index    'Completa con ceros a la izquierda.
   'Case 0, 1                           'Cambiar (añadir índices).
   '   If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
   '      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
   '   End If
   'End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1            'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
    Case 0, 1
       If txtDato(tnIndex).Text = "" Then
          lblDatoDeta(tnIndex).Caption = ""
          Exit Function
       End If
       With porstCOCta
          .MoveFirst
          .Find "CodCta='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
             MsgBox TEXT_8006, vbExclamation
             ppAyuDet = True
          Else
             lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
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
Public Sub Llena_Temporal()

    'jp
    Dim aImporte(2, 4) As Double
    Dim nSaldo As Double, nSaldoAcu As Double
    Dim nPresupuesto As Double, nPresupuestoAcu As Double
    
    pocnnMain.Execute "DELETE FROM COTmpRpt WHERE codemp='" & gsCodEmp & "' AND UsrCre='" & gsCodUsr & "' AND NomRpt='rptRCtlPsp'"
    porstMRpRs.Open
    If porstMRp.RecordCount > 0 Then
        porstMRp.MoveFirst
        Do While Not porstMRp.EOF
           nPresupuesto = porstMRp.Fields!cPreMes
           nPresupuestoAcu = porstMRp.Fields!cPreAcu
           If porstMRp.Fields!cOrden = "A" Then
             nSaldo = CDec(IIf(Not IsNull(porstMRp.Fields!clmpMesHaber), porstMRp.Fields!clmpMesHaber, 0) - IIf(Not IsNull(porstMRp.Fields!clmpMesDebe), porstMRp.Fields!clmpMesDebe, 0))
             nSaldoAcu = CDec(IIf(Not IsNull(porstMRp.Fields!clmpAcuHaber), porstMRp.Fields!clmpAcuHaber, 0) - IIf(Not IsNull(porstMRp.Fields!clmpAcuDebe), porstMRp.Fields!clmpAcuDebe, 0))
             aImporte(1, 1) = aImporte(1, 1) + nSaldo
             aImporte(1, 2) = aImporte(1, 2) + nSaldoAcu
             aImporte(1, 3) = aImporte(1, 3) + nPresupuesto
             aImporte(1, 4) = aImporte(1, 4) + nPresupuestoAcu
           Else
             nSaldo = CDec(IIf(Not IsNull(porstMRp.Fields!clmpMesDebe), porstMRp.Fields!clmpMesDebe, 0) - IIf(Not IsNull(porstMRp.Fields!clmpMesHaber), porstMRp.Fields!clmpMesHaber, 0))
             nSaldoAcu = CDec(IIf(Not IsNull(porstMRp.Fields!clmpAcuDebe), porstMRp.Fields!clmpAcuDebe, 0) - IIf(Not IsNull(porstMRp.Fields!clmpAcuHaber), porstMRp.Fields!clmpAcuHaber, 0))
             aImporte(2, 1) = aImporte(2, 1) + nSaldo
             aImporte(2, 2) = aImporte(2, 2) + nSaldoAcu
             aImporte(2, 3) = aImporte(2, 3) + nPresupuesto
             aImporte(2, 4) = aImporte(2, 4) + nPresupuestoAcu
           End If
           If (nSaldo <> 0 Or nSaldoAcu <> 0 Or nPresupuesto <> 0 Or nPresupuestoAcu <> 0) Then
                pocnnMain.BeginTrans
                porstMRpRs.AddNew
                porstMRpRs.Fields!codemp = gsCodEmp
                porstMRpRs.Fields!pdoano = gsAnoAct
                porstMRpRs.Fields!UsrCre = gsCodUsr
                porstMRpRs.Fields!NomRpt = "rptRCtlPsp"
                porstMRpRs.Fields!CodCta = porstMRp.Fields!CodCta
                porstMRpRs.Fields!codcco = porstMRp.Fields!codcco
                porstMRpRs.Fields!detcta = IIf(IsNull(porstMRp.Fields!detcta), "", porstMRp.Fields!detcta)
                porstMRpRs.Fields!ordrep_1 = porstMRp.Fields!cOrden
                porstMRpRs.Fields!DetOrdRep_1 = porstMRp.Fields!cTitulo
                porstMRpRs.Fields!OrdRep_2 = porstMRp.Fields!cTitulo2
                porstCOCta.MoveFirst
                porstCOCta.Find "CodCta='" & Left(porstMRp.Fields!CodCta, 2) & "'"
                If Not porstCOCta.EOF Then
                    porstMRpRs.Fields!DetOrdRep_2 = porstCOCta.Fields!detcta
                End If
                porstMRpRs.Fields!ImpSdo_Mes = nSaldo
                porstMRpRs.Fields!ImpSdo_Acu = nSaldoAcu
                porstMRpRs.Fields!PspMes = nPresupuesto
                porstMRpRs.Fields!PspAcu = nPresupuestoAcu
                porstMRpRs.Update
                pocnnMain.CommitTrans
            End If
            porstMRp.MoveNext
        Loop
        ' Actualizo los importes totales de cada rubro
        pocnnMain.BeginTrans
        porstMRpRs.Fields!NumCol0 = Round(aImporte(1, 1) - aImporte(2, 1), 2)
        porstMRpRs.Fields!NumCol1 = Round(aImporte(1, 2) - aImporte(2, 2), 2)
        porstMRpRs.Fields!NumCol2 = Round(aImporte(1, 3) - aImporte(2, 3), 2)
        porstMRpRs.Fields!numCol3 = Round(aImporte(1, 4) - aImporte(2, 4), 2)
        porstMRpRs.Update
        pocnnMain.CommitTrans
    End If
    'jp

End Sub

Private Sub txtllave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   'If pbValidada Then txtDato(0).SetFocus
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBusx Index
   End If
End Sub

Private Sub ppAyuBusx(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "length(codcco)=2 ", "", 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
      txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
      lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub


