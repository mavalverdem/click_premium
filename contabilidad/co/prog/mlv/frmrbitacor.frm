VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRBitacor 
   Caption         =   "[título]"
   ClientHeight    =   4590
   ClientLeft      =   1620
   ClientTop       =   1410
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5115
   Begin VB.Frame FrameRep 
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   3160
      Width           =   2655
      Begin MSComCtl2.DTPicker dtpDato 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   26
         Top             =   150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   140050433
         CurrentDate     =   37102
      End
      Begin MSComCtl2.DTPicker dtpDato 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   140050433
         CurrentDate     =   37102
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   " Transacción "
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   60
      Width           =   5000
      Begin VB.OptionButton optReporte 
         Caption         =   "Auxiliares"
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   31
         Top             =   720
         Width           =   2100
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Plan de Cuentas"
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   30
         Top             =   480
         Width           =   2100
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Registro de Diarios"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   2100
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Registro de Honorarios"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   2100
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Registro de Ventas"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   2100
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Registro de Compras"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2820
      TabIndex        =   16
      Top             =   3240
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   315
         Width           =   915
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   2115
      Left            =   0
      TabIndex        =   12
      Top             =   1050
      Width           =   5000
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   4455
         Picture         =   "frmrbitacor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
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
         Top             =   480
         Width           =   570
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1650
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
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1650
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   4455
         Picture         =   "frmrbitacor.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   4455
         Picture         =   "frmrbitacor.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   4750
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de Documento"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1485
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
         Height          =   300
         Index           =   0
         Left            =   675
         TabIndex        =   23
         Top             =   480
         Width           =   3795
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
         Height          =   300
         Index           =   2
         Left            =   1755
         TabIndex        =   15
         Top             =   1680
         Width           =   2715
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
         Height          =   300
         Index           =   1
         Left            =   1755
         TabIndex        =   14
         Top             =   1320
         Width           =   2715
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Documento"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1650
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   5115
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4050
      Width           =   5115
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
         Height          =   495
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
         Height          =   495
         Left            =   3720
         Picture         =   "frmrbitacor.frx":04FE
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
         Height          =   495
         Index           =   0
         Left            =   0
         Picture         =   "frmrbitacor.frx":0648
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
         Height          =   495
         Index           =   1
         Left            =   1245
         Picture         =   "frmrbitacor.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRBitacor"
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
Private porstFiltro As ADODB.Recordset
Private porstRango As ADODB.Recordset
Private n_Reporte As Integer

Public reporte As ADODB.Recordset
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

   FrameRep.Visible = False
   dtpDato(0).Value = "01/01/" & Year(Date)
   dtpDato(1).Value = Date
   

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstFiltro = New ADODB.Recordset
   Set porstRango = New ADODB.Recordset
   
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
   With porstFiltro
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
   End With
   With porstRango
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
   End With
 ']

 '[Parámetros.                         'Cambiar.
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Tipo de Documento :", "Numero de Documento :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Type of Document :", "Number of Document :")
  Next nElemento
  fraTipo.Caption = Choose(gsIdioma, "Transacción", "Transaction")
  optReporte(0).Caption = Choose(gsIdioma, "Registro de Compras", "Purchase Register")
  optReporte(1).Caption = Choose(gsIdioma, "Registro de Ventas", "Sales Register")
  optReporte(2).Caption = Choose(gsIdioma, "Registro de Honorarios", "Feeds Register")
  optReporte(3).Caption = Choose(gsIdioma, "Registro de Diarios", "Journals Register")
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
 
 '[Datos predeterminados.              'Cambiar.
  optReporte(0).Value = True
  
  'Otros.
   
  'Características de impresión.
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
   porstRango.Close
   pocnnMain.Close
   Set porstRango = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim s_Sentencia As String, s_Temporal As String
  
  If optReporte(4).Value = True Then
  
   Set reporte = New ADODB.Recordset

   With reporte
      .ActiveConnection = pocnnMain
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
     
   With reporte
   If .State = adStateOpen Then .Close
    .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta, "
    .Source = .Source & "(CASE TpoCta WHEN " & TPOCTA_TIT & " THEN '" & TPOCTA_TIT_TXT & "' WHEN " & TPOCTA_TRA & " THEN '" & TPOCTA_TRA_TXT & "' END) AS cTpoCta, "
    .Source = .Source & "(CASE TpoSdo WHEN '" & TPOSDO_INV & "' THEN 'Inv.' "
    .Source = .Source & "WHEN '" & TPOSDO_RES & "' THEN 'Res.' "
    .Source = .Source & "WHEN '" & TPOSDO_FUN & "' THEN 'Func.' "
    .Source = .Source & "WHEN '" & TPOSDO_NAT & "' THEN 'Nat.' "
    .Source = .Source & "WHEN '" & TPOSDO_AMB & "' THEN 'F/N' END) AS cTpoSdo, "
    .Source = .Source & "(CASE IndDoc WHEN " & INDDOC_INA & " THEN 'No' WHEN " & INDDOC_ACT & " THEN 'Si' END) AS cIndDoc, "
    .Source = .Source & "(CASE IndCCo WHEN " & INDCCO_INA & " THEN 'No' WHEN " & INDCCO_ACT & " THEN 'Si' END) AS cIndCCo, "
    .Source = .Source & "(CASE IndPsp WHEN " & INDPSP_INA & " THEN 'No' WHEN " & INDPSP_ACT & " THEN 'Si' END) AS cIndPsp, "
    .Source = .Source & "(CASE IndMoe WHEN " & INDMOE_INA & " THEN 'No' WHEN " & INDMOE_ACT & " THEN 'Si' END) AS cIndMoe, "
    .Source = .Source & "CodCta_Dst_Deb, CodCta_Dst_Hab, "
    .Source = .Source & "(CASE IndAjd WHEN " & INDAJD_INA & " THEN 'No' WHEN " & INDAJD_ACT & " THEN 'Si' END) AS cIndAjd, "
    .Source = .Source & "(CASE IndAjd WHEN " & INDAJD_ACT & " THEN "
    .Source = .Source & "(CASE TpoAjD WHEN '" & TPOANL_CTA & "' THEN '" & TPOANL_CTA_TXT & "' WHEN '" & TPOANL_AUX & "' THEN '" & TPOANL_AUX_TXT & "' "
    .Source = .Source & "WHEN '" & TPOANL_DOC & "' THEN '" & TPOANL_DOC_TXT & "'  END) END) AS cTpoAjD, "
    .Source = .Source & "(CASE IndAjd WHEN " & INDAJD_ACT & " THEN "
    .Source = .Source & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & TPOMON_NAC_TXT_0 & "'  WHEN '" & TPOMON_EXT & "' THEN '" & TPOMON_EXT_TXT_0 & "'  END) END) AS cTpoMon, "
    .Source = .Source & "(CASE IndAjd WHEN " & INDAJD_ACT & " THEN "
    .Source = .Source & "(CASE TpoTCb WHEN '" & TPOTCB_CPR & "' THEN '" & TPOTCB_CPR_TXT & "'  WHEN '" & TPOTCB_VTA & "' THEN '" & TPOTCB_VTA_TXT & "' END) END) AS cTpoTCb, "
    .Source = .Source & "(CASE IndAjd WHEN " & INDAJD_ACT & " THEN CodCta_Ajd_Deb END) AS  cCodCta_Ajd_Deb, "
    .Source = .Source & "(CASE IndAjd WHEN " & INDAJD_ACT & " THEN CodCta_Ajd_Hab END) AS cCodCta_Ajd_Hab, "
    .Source = .Source & "(CASE EstCta WHEN '" & ESTCTA_ACT & "' THEN '" & ESTCTA_ACT_TXT & "' WHEN '" & ESTCTA_INA & "' THEN '" & ESTCTA_INA_TXT & "' END) AS cEstCta, "
    .Source = .Source & "(CASE NatCta WHEN " & NATCTA_DEU & " THEN '" & NATCTA_DEU_TXT & "' WHEN " & NATCTA_ACR & " THEN '" & NATCTA_ACR_TXT & "' END) AS cNatuCta "
    .Source = .Source & "FROM COCta "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' AND "
    .Source = .Source & " fyhcre Between '" & Format(dtpDato(0).Value, "yyyy-mm-dd hh:mm:ss") & "' and '" & Format(dtpDato(1).Value, "yyyy-mm-dd hh:mm:ss") & "'"
    .Source = .Source & " ORDER BY CodCta"
    .Open
   End With
      
   gpEncabezadoRpt frmMain.rptMain, "Listado de Cuentas Contables Ingresadas desde " & dtpDato(0).Value & " al " & dtpDato(1).Value, Date, True, False, reporte
   With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLcta.rpt"
      '.MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
   End With

  Exit Sub
  
  End If
  
  If optReporte(5).Value = True Then
  
   Set reporte = New ADODB.Recordset

   With reporte
      .ActiveConnection = pocnnMain
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
     
   With reporte
   If .State = adStateOpen Then .Close
    .Source = "SELECT TGAux.CodAux, TGAux.RazAux, TGAux.RucAux, "
    .Source = .Source & "(Case TGAux.TpoPer when 'J' then 'Juridica' when 'N' then 'Natural' End) AS cTpoPer, "
    .Source = .Source & "TGAuxNat.NomAux, TGAuxNat.ApePatAux, TGAuxNat.ApeMatAux, "
    .Source = .Source & "(Case TGAux.IndCli  when '1' then 'X'  End) as cIndCli, "
    .Source = .Source & "(Case TGAux.IndPrv  when '1' then 'X' End) as cIndPrv, "
    .Source = .Source & "(Case TGAux.IndOtr  when '1' then 'X' End) as cIndOtr, "
    .Source = .Source & "(Case TGAux.EstAux when 'A' Then 'Activo' Else 'Inactivo' End) as vEstAux "
    .Source = .Source & "FROM TGAux "
    .Source = .Source & "LEFT JOIN TGAuxNat ON TGaux.codemp=TGAuxNat.codemp AND TGAux.CodAux = TGAuxNat.CodAux "
    .Source = .Source & "WHERE TGAux.codemp='" & gsCodEmp & "' and "
    .Source = .Source & " tgaux.fyhcre Between '" & Format(dtpDato(0).Value, "yyyy-mm-dd hh:mm:ss") & "' and '" & Format(dtpDato(1).Value, "yyyy-mm-dd hh:mm:ss") & "'"
    .Source = .Source & " ORDER BY TGAux.CodAux"
    .Open
   End With
      
   gpEncabezadoRpt frmMain.rptMain, "Listado de Auxiliares Ingresadas desde " & dtpDato(0).Value & " al " & dtpDato(1).Value, Date, True, False, reporte
   With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLAux.rpt"
      '.MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
   End With

  Exit Sub
  
  End If
  
  ppHabilitacion False
  s_Temporal = Choose(n_Reporte + 1, "trptRBitaCpr", "trptRBitaVta", "trptRBitaHpr", "trptRBitaDro")
  s_Temporal = ps_Prefijo & s_Temporal
  
  s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS " & s_Temporal & " ", "")
  If n_Reporte = 0 Then
    s_Sentencia = s_Sentencia & "SELECT a.CodAux, a.CodTDc, d.AbvTDc, " & Choose(gsIdioma, "d.DetTDc", "d.DetTDcx") & " AS DetTDc, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, '-', a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    s_Sentencia = s_Sentencia & "a.FeEDoc, a.GloDoc, (CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpOGr_MN + a.ImpOGN_MN + a.ImpONG_MN + a.ImpExo_MN, 2) AS nBase_MN, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpIGV_MN + a.ImpISC_MN + a.ImpOIm_MN, 2) AS nImpuesto_MN, a.ImpTot_MN, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpOGr_ME + a.ImpOGN_ME + a.ImpONG_ME + a.ImpExo_ME, 2) AS nBase_ME, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpIGV_ME + a.ImpISC_ME + a.ImpOIm_ME, 2) AS nImpuesto_ME, a.ImpTot_ME, "
    s_Sentencia = s_Sentencia & "a.UsrCre AS UsrCre1, a.FyHCre AS FyHCre1, a.UsrMdf AS UsrMdf1, a.FyHMdf AS FyHMdf1, "
    s_Sentencia = s_Sentencia & "b.TpoCnc, b.Orden, b.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS DetCta, b.GloDet, b.ImpCta_MN, b.ImpCta_ME, "
    s_Sentencia = s_Sentencia & "b.UsrCre AS UsrCre2, b.FyHCre AS FyHCre2, b.UsrMdf AS UsrMdf2, b.FyHMdf AS FyHMdf2, "
    s_Sentencia = s_Sentencia & "c.CodCCo, " & Choose(gsIdioma, "f.DetCCo", "f.DetCCox") & " AS DetCCo, c.ImpCCo_MN, c.ImpCCo_ME, "
    s_Sentencia = s_Sentencia & "c.UsrCre AS UsrCre3, c.FyHCre AS FyHCre3, c.UsrMdf AS UsrMdf3, c.FyHMdf AS FyHMdf3 "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & s_Temporal & " ", "")
    s_Sentencia = s_Sentencia & "FROM (((((CoCprDoc a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCprDocCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCprDocCCo c ON b.codemp=c.codemp AND b.pdoano=c.pdoano AND b.CodAux=c.CodAux AND b.CodTDc=c.CodTDc AND b.SerDoc=c.SerDoc AND b.NroDoc=c.NroDoc AND b.TpoCnc=c.TpoCnc AND b.Orden=c.Orden AND b.CodCta=c.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta e ON b.codemp=e.codemp AND b.pdoano=e.pdoano AND b.CodCta=e.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo f ON c.codemp=f.codemp AND c.pdoano=f.pdoano AND c.CodCCo=f.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc) "
  ElseIf n_Reporte = 1 Then
    s_Sentencia = s_Sentencia & "SELECT a.CodAux, a.CodTDc, d.AbvTDc, " & Choose(gsIdioma, "d.DetTDc", "d.DetTDcx") & " AS DetTDc, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, '-', a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    s_Sentencia = s_Sentencia & "a.FeEDoc, a.GloDoc, (CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpOGr_MN + a.ImpExp_MN + a.ImpExo_MN, 2) AS nBase_MN, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpIGV_MN + a.ImpISC_MN + a.ImpOIm_MN, 2) As nImpuesto_MN, a.ImpTot_MN, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpOGr_ME + a.ImpExp_ME + a.ImpExo_ME, 2) AS nBase_ME, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpIGV_ME + a.ImpISC_ME + a.ImpOIm_ME, 2) As nImpuesto_ME, a.ImpTot_ME, "
    s_Sentencia = s_Sentencia & "a.UsrCre AS UsrCre1, a.FyHCre AS FyHCre1, a.UsrMdf AS UsrMdf1, a.FyHMdf AS FyHMdf1, "
    s_Sentencia = s_Sentencia & "b.TpoCnc, b.Orden, b.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS DetCta, b.GloDet0 AS GloDet, b.ImpCta_MN, b.ImpCta_ME, "
    s_Sentencia = s_Sentencia & "b.UsrCre AS UsrCre2, b.FyHCre AS FyHCre2, b.UsrMdf AS UsrMdf2, b.FyHMdf AS FyHMdf2, "
    s_Sentencia = s_Sentencia & "c.CodCCo, " & Choose(gsIdioma, "f.DetCCo", "f.DetCCox") & " AS DetCCo, c.ImpCCo_MN, c.ImpCCo_ME, "
    s_Sentencia = s_Sentencia & "c.UsrCre AS UsrCre3, c.FyHCre AS FyHCre3, c.UsrMdf AS UsrMdf3, c.FyHMdf AS FyHMdf3 "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & s_Temporal & " ", "")
    s_Sentencia = s_Sentencia & "FROM (((((CoVtaDoc a LEFT JOIN CoVtaDocCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoVtaDocCCo c ON b.codemp=c.codemp AND b.pdoano=c.pdoano AND b.CodTDc=c.CodTDc AND b.SerDoc=c.SerDoc AND b.NroDoc=c.NroDoc AND b.TpoCnc=c.TpoCnc AND b.Orden=c.Orden AND b.CodCta=c.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta e ON b.codemp=e.codemp AND b.pdoano=e.pdoano AND b.CodCta=e.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo f ON c.codemp=f.codemp AND c.pdoano=f.pdoano AND c.CodCCo=f.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc) "
  ElseIf n_Reporte = 2 Then
    s_Sentencia = s_Sentencia & "SELECT a.CodAux, '" & CODTDC_HPR & "' AS CodTDc, '" & lblDatoDeta(0).Caption & "' AS DetTDc, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, '-', a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cDocumento, "
    s_Sentencia = s_Sentencia & "a.FeEDoc, a.GloDoc, (CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
    s_Sentencia = s_Sentencia & "a.ImpBru_MN AS nBase_MN, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpIR4_MN + a.ImpIES_MN + a.ImpORt_MN, 2) AS nImpuesto_MN, a.ImpNet_MN, "
    s_Sentencia = s_Sentencia & "a.ImpBru_ME AS nBase_ME, "
    s_Sentencia = s_Sentencia & "ROUND(a.ImpIR4_ME + a.ImpIES_ME + a.ImpORt_ME, 2) AS nImpuesto_ME, a.ImpNet_ME, "
    s_Sentencia = s_Sentencia & "a.UsrCre AS UsrCre1, a.FyHCre AS FyHCre1, a.UsrMdf AS UsrMdf1, a.FyHMdf AS FyHMdf1, "
    s_Sentencia = s_Sentencia & "b.TpoCnc, b.Orden, b.CodCta, " & Choose(gsIdioma, "e.DetCta", "e.DetCtax") & " AS DetCta, b.GloDet, b.ImpCta_MN, b.ImpCta_ME, "
    s_Sentencia = s_Sentencia & "b.UsrCre AS UsrCre2, b.FyHCre AS FyHCre2, b.UsrMdf AS UsrMdf2, b.FyHMdf AS FyHMdf2, "
    s_Sentencia = s_Sentencia & "c.CodCCo, " & Choose(gsIdioma, "f.DetCCo", "f.DetCCox") & " AS DetCCo, c.ImpCCo_MN, c.ImpCCo_ME, "
    s_Sentencia = s_Sentencia & "c.UsrCre AS UsrCre3, c.FyHCre AS FyHCre3, c.UsrMdf AS UsrMdf3, c.FyHMdf AS FyHMdf3 "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & s_Temporal & " ", "")
    s_Sentencia = s_Sentencia & "FROM ((((CoHprDoc a  "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoHprDocCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoHprDocCCo c ON b.codemp=c.codemp AND b.pdoano=c.pdoano AND b.CodAux=c.CodAux AND b.SerDoc=c.SerDoc AND b.NroDoc=c.NroDoc AND b.TpoCnc=c.TpoCnc AND b.Orden=c.Orden AND b.CodCta=c.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta e ON b.codemp=e.codemp AND b.pdoano=e.pdoano AND b.CodCta=e.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo f ON c.codemp=f.codemp AND c.pdoano=f.pdoano AND c.CodCCo=f.CodCCo) "
  ElseIf n_Reporte = 3 Then
    s_Sentencia = s_Sentencia & "SELECT a.CodDro, c.DetDro, a.NroCpb, " & Choose(gsIdioma, "a.GloCpb", "a.GloCpbx") & " AS GloCpb, a.UsrCre AS UsrCre1, a.FyHCre AS FyHCre1, a.UsrMdf AS UsrMdf1, a.FyHMdf AS FyHMdf1, "
    s_Sentencia = s_Sentencia & "b.NroIte, b.CodCta, b.CodCCo, b.CodAux, " & Choose(gsIdioma, "b.GloIte", "b.GloItex") & " AS GloIte, d.AbvTDc, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.SerDoc, '-', b.NroDoc)", "(b.SerDoc+'-'+b.NroDoc)") & " AS cDocumento, "
    s_Sentencia = s_Sentencia & "(CASE b.TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
    s_Sentencia = s_Sentencia & "(CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END) AS cCargoMN, (CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END) AS cAbonoMN, "
    s_Sentencia = s_Sentencia & "(CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END) AS cCargoME, (CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END) AS cAbonoME, "
    s_Sentencia = s_Sentencia & "b.UsrCre AS UsrCre2, b.FyHCre AS FyHCre2, b.UsrMdf AS UsrMdf2, b.FyHMdf AS FyHMdf2 "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & s_Temporal & " ", "")
    s_Sentencia = s_Sentencia & "FROM (((CoCpbDet b "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCpbCab a ON b.codemp=a.codemp AND b.pdoano=a.pdoano AND b.MesPvs=a.MesPvs AND b.CodDro=a.CodDro AND b.NroCpb=a.NroCpb) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoDro c ON b.codemp=c.codemp AND b.pdoano=c.pdoano AND b.CodDro=c.CodDro) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc d ON b.codemp=d.codemp AND b.CodTDc=d.CodTDc) "
      
  End If
  s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND a.pdoano='" & gsAnoAct & "' "
  s_Sentencia = s_Sentencia & "AND a.MesPvs='" & gsMesAct & "' "
  If n_Reporte <> 2 Then
    s_Sentencia = s_Sentencia & "AND " & IIf(n_Reporte = 3, "a.CodDro", "a.CodTDc") & "='" & txtDato(0).Text & "' "
  End If
  s_Sentencia = s_Sentencia & "AND " & IIf(n_Reporte = 3, "a.NroCpb", IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, '-', a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)")) & " BETWEEN '" & txtDato(1).Text & "' AND '" & txtDato(2).Text & "' "
  s_Sentencia = s_Sentencia & "ORDER BY " & Choose(n_Reporte + 1, "a.CodAux, cDocumento, b.TpoCnc, b.Orden", "cDocumento, b.TpoCnc, b.Orden", "a.CodAux, cDocumento, b.TpoCnc, b.Orden", "a.NroCpb, b.NroIte")
   
  'Genero la tabla temporal de impresion
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & s_Temporal, "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='" & s_Temporal & "') DROP TABLE " & s_Temporal)
  pocnnMain.Execute s_Sentencia
  
  ' Actualizo los nombres de usuarios cifrados
  ppActualizoUsr s_Temporal
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT * FROM " & s_Temporal & " ORDER BY " & Choose(n_Reporte + 1, "CodAux, cDocumento, TpoCnc, Orden", "cDocumento, TpoCnc, Orden", "CodAux, cDocumento, TpoCnc, Orden", "NroCpb, NroIte")
    .Open
  End With
   
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " - " & optReporte(n_Reporte).Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & Choose(n_Reporte + 1, "rptRBitaCpr.rpt", "rptRBitaVta.rpt", "rptRBitaHpr.rpt", "rptRBitaDro.rpt")
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .WindowState = crptMaximized
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
    .DataRecordSet = porstMRp
    .LoadReport gsRutRpt & "rptLDro.mrp"
    
    Call gpEncabezadoMRp(MRViewer, Me.Caption, udFecha, True)
    '[Parámetros adicionales.
    '         .Parameters("pTipoFecha") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
  ' Elimino la tabla temporal de impresion
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & s_Temporal, "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='" & s_Temporal & "') DROP TABLE " & s_Temporal)
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
Private Sub optReporte_Click(Index As Integer)
    
  Select Case Index
    Case 4, 5
    FrameRep.Visible = True
    fraRangos.Visible = False
    optTipoImpresion(1).Value = True
    optTipoImpresion(0).Visible = False
    Exit Sub
    Case Else
    FrameRep.Visible = False
    fraRangos.Visible = True
    optTipoImpresion(0).Value = True
    optTipoImpresion(0).Visible = True
  End Select
    
  n_Reporte = Index
  If gsIdioma = NvlUsr_Sup Then
    lblTexto(0).Caption = Choose(n_Reporte + 1, "Tipo de Documento", "Tipo de Documento", "Tipo de Documento", "Diario")
    lblTexto(1).Caption = Choose(n_Reporte + 1, "Documentos de Compra", "Documentos de Venta", "Documentos de Honorario", "Comprobante de Diario")
  Else
    lblTexto(0).Caption = Choose(n_Reporte + 1, "Type of Document", "Type of Document", "Type of Document", "Journal")
    lblTexto(1).Caption = Choose(n_Reporte + 1, "Documents of Purchase", "Documents of Sale", "Documents of Feed", "Voucher of Journal")
  End If
  
 '[ Activo el recorset de acuerdo al caso de
  With porstFiltro
    If .State = adStateOpen Then .Close
    .Source = IIf(n_Reporte = 3, "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro FROM CoDro WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro) > 2 ORDER BY CodDro", "SELECT CodTDc, DetTDc FROM TGTDc WHERE codemp='" & gsCodEmp & "' ORDER BY CodTDc")
    .Open
  End With
  txtDato(0).DataField = IIf(n_Reporte = 3, "CodDro", "CodTDc")
  txtDato(0).MaxLength = porstFiltro.Fields(txtDato(0).DataField).DefinedSize
  If n_Reporte = 2 Then
    txtDato(0).Text = CODTDC_HPR
  Else
    txtDato(0).Text = porstFiltro.Fields(txtDato(0).DataField)
  End If
  txtDato(0).Enabled = (n_Reporte <> 2)
  cmdDatoAyud(0).Enabled = (n_Reporte <> 2)
  
 'Busca detalle de códigos            '(habilitar/deshabilitar).
  If txtDato(0).Text <> "" Then ppAyuDet 0
  ppRango
  
End Sub
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

Private Sub txtDato_LostFocus(Index As Integer)
  If Index = 0 Then ppRango
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index    'Completa con ceros a la izquierda.
   Case 0                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   Case 1, 2                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                          'Cambiar (añadir índices).
      If n_Reporte = 3 Then
         modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro) > 2", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      Else
         modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      End If
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1, 2                          'Cambiar (añadir índices).
      If n_Reporte = 0 Then
        modAyuBus.Doc_Cpr "CodTDc='" & txtDato(0).Text & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      ElseIf n_Reporte = 1 Then
        modAyuBus.Doc_Vta "CodTDc='" & txtDato(0).Text & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      ElseIf n_Reporte = 2 Then
        modAyuBus.Doc_Hpr "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      ElseIf n_Reporte = 3 Then
        modAyuBus.Cpb_Dro "CodDro='" & txtDato(0).Text & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      End If
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
      With porstFiltro
         .MoveFirst
         .Find IIf(n_Reporte = 3, "CodDro", "CodTDc") & "='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & porstFiltro(IIf(n_Reporte = 3, "DetDro", "DetTDc"))
         End If
      End With
   Case 1, 2
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstRango
         .MoveFirst
         .Find IIf(n_Reporte = 3, "NroCpb", "cDocumento") & "='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & porstRango(IIf(n_Reporte = 3, "GloCpb", "GloDoc"))
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

'[
Private Sub ppRango()
  Dim s_Sentencia As String
  Dim n_Index As Integer
  
  If n_Reporte = 3 Then
    s_Sentencia = "SELECT NroCpb, " & Choose(gsIdioma, "GloCpb", "GloCpbx") & " AS GloCpb "
    s_Sentencia = s_Sentencia & "FROM CoCpbCab "
    s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND MesPvs='" & gsMesAct & "' "
    s_Sentencia = s_Sentencia & "AND CodDro='" & txtDato(0).Text & "' "
    s_Sentencia = s_Sentencia & "ORDER BY NroCpb"
  Else
    s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc, '-', NroDoc)", "(SerDoc+'-'+NroDoc)") & " AS cDocumento, GloDoc, SerDoc, NroDoc "
    s_Sentencia = s_Sentencia & "FROM " & Choose(n_Reporte + 1, "CoCprDoc", "CoVtaDoc", "CoHprDoc") & " "
    s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND MesPvs='" & gsMesAct & "' "
    If n_Reporte <> 2 Then
      s_Sentencia = s_Sentencia & "AND CodTDc='" & txtDato(0).Text & "' "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY SerDoc, NroDoc"
  End If
  With porstRango
    If .State = adStateOpen Then .Close
    .Source = s_Sentencia
    .Open
  End With
  With txtDato
    For n_Index = 1 To 2
      .Item(n_Index).DataField = IIf(n_Reporte = "3", "NroCpb", "cDocumento")
      .Item(n_Index).MaxLength = porstRango.Fields(.Item(n_Index).DataField).DefinedSize
      .Item(n_Index).Text = ""
      lblDatoDeta(n_Index) = ""
    Next
  End With
  'Límites de rangos.
   If Not (porstRango.EOF And porstRango.BOF) Then
     With porstRango
       .MoveLast
       txtDato(2).Text = .Fields(txtDato(2).DataField)
       .MoveFirst
       txtDato(1).Text = .Fields(txtDato(2).DataField)
     End With
   End If
  
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2

End Sub
Private Sub ppActualizoUsr(s_Temporal)
  Dim porstTemporal As ADODB.Recordset
  Dim nContador As Integer
  Dim s_Usuario As String
  Dim sSentencia As String
  
  Set porstTemporal = New ADODB.Recordset

  With porstTemporal
    .ActiveConnection = pocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = "SELECT DISTINCT UsrCre1, UsrCre2, " & IIf(n_Reporte <> 3, "UsrCre3, ", " ")
    .Source = .Source & "UsrMdf1, UsrMdf2" & IIf(n_Reporte <> 3, ", UsrMdf3 ", " ")
    .Source = .Source & "FROM " & s_Temporal
    .Open
    If Not (.EOF And .BOF) Then .MoveFirst
  End With

  pocnnMain.BeginTrans    '[ INICIA TRANSACCION ]
  Do While Not porstTemporal.EOF
    For nContador = 1 To Choose(n_Reporte + 1, 3, 3, 3, 2)
      s_Usuario = IIf(IsNull(porstTemporal.Fields("UsrCre" & Trim$(nContador))), "", porstTemporal.Fields("UsrCre" & Trim$(nContador)))
      If s_Usuario <> "" Then
        sSentencia = "UPDATE " & s_Temporal & " "
        sSentencia = sSentencia & "SET " & "UsrCre" & Trim$(nContador) & "='" & gfEnmasc(s_Usuario) & "' "
        sSentencia = sSentencia & "WHERE " & "UsrCre" & Trim$(nContador) & "='" & porstTemporal("UsrCre" & Format(nContador, "0")) & "' "
        pocnnMain.Execute sSentencia
      End If
      s_Usuario = IIf(IsNull(porstTemporal.Fields("UsrMdf" & Trim$(nContador))), "", porstTemporal.Fields("UsrMdf" & Trim$(nContador)))
      If s_Usuario <> "" Then
        sSentencia = "UPDATE " & s_Temporal & " "
        sSentencia = sSentencia & "SET " & "UsrMdf" & Trim$(nContador) & "='" & gfEnmasc(s_Usuario) & "' "
        sSentencia = sSentencia & "WHERE " & "UsrMdf" & Trim$(nContador) & "='" & porstTemporal("UsrMdf" & Format(nContador, "0")) & "' "
        pocnnMain.Execute sSentencia
      End If
    Next
    porstTemporal.MoveNext
  Loop
  pocnnMain.CommitTrans
  porstTemporal.Close
  Set porstTemporal = Nothing
  
End Sub
']
