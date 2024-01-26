VERSION 5.00
Begin VB.Form frmREFiCCo 
   Caption         =   "[título]"
   ClientHeight    =   4905
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkpresupuesto 
      Caption         =   "Presupuesto"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   2520
      TabIndex        =   37
      Top             =   1800
      Width           =   1980
   End
   Begin VB.CheckBox chkDivisoria 
      Caption         =   "Divisionarias"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   2520
      TabIndex        =   36
      Top             =   1440
      Width           =   1980
   End
   Begin VB.CheckBox chkRango 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1455
      TabIndex        =   17
      Top             =   2880
      Width           =   180
   End
   Begin VB.Frame fraRngPeriodo 
      Caption         =   " Rango Periodos "
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   60
      TabIndex        =   16
      Top             =   2880
      Width           =   4215
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   23
         Text            =   "Mes Final"
         Top             =   645
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         ItemData        =   "frmREFiCCo.frx":0000
         Left            =   2310
         List            =   "frmREFiCCo.frx":0002
         TabIndex        =   20
         Text            =   "Mes Inicio"
         Top             =   300
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   855
         TabIndex        =   22
         Text            =   "Año Final"
         Top             =   645
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   19
         Text            =   "Año Inicio"
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.Frame fraFormato 
      Caption         =   " Formato "
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   60
      TabIndex        =   13
      Top             =   2070
      Width           =   4695
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
         TabIndex        =   14
         Top             =   255
         Width           =   735
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   4290
         Picture         =   "frmREFiCCo.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   270
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
         Index           =   2
         Left            =   840
         TabIndex        =   15
         Top             =   255
         Width           =   3450
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   1830
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   24
      Top             =   3330
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   315
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   26
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCCostos 
      Caption         =   "C.C&ostos"
      Height          =   375
      Left            =   5760
      TabIndex        =   32
      Top             =   2385
      Width           =   1215
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      Begin VB.CheckBox Todos 
         Alignment       =   1  'Right Justify
         Caption         =   "Todas las Cuentas"
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton optFormato 
         Caption         =   "Cen. Costos :"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1710
         TabIndex        =   2
         Top             =   220
         Value           =   -1  'True
         Width           =   1300
      End
      Begin VB.OptionButton optFormato 
         Caption         =   "Cuentas :"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   220
         Width           =   1300
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
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6660
         Picture         =   "frmREFiCCo.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   855
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
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6660
         Picture         =   "frmREFiCCo.frx":0358
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   495
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
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   5595
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
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   5595
      End
   End
   Begin VB.Frame fraAlcance 
      Caption         =   "Alcance"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   15
      TabIndex        =   7
      Top             =   1350
      Width           =   2445
      Begin VB.OptionButton optAlcance 
         Caption         =   "al mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optAlcance 
         Caption         =   "del mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   9
         Top             =   255
         Width           =   1080
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1395
      Width           =   1260
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7035
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4365
      Width           =   7035
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
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
         Left            =   3600
         Picture         =   "frmREFiCCo.frx":0502
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         Width           =   1125
      End
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
         TabIndex        =   29
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
         Picture         =   "frmREFiCCo.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Picture         =   "frmREFiCCo.frx":074E
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   5880
         Picture         =   "frmREFiCCo.frx":0C80
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label mensaje 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   4965
      TabIndex        =   10
      Top             =   1440
      Width           =   705
   End
End
Attribute VB_Name = "frmREFiCCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MRViewer As MRViewerObject
Attribute MRViewer.VB_VarHelpID = -1
Dim cnn As ADODB.Connection

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
Private porstMRpRs          As ADODB.Recordset
Private porstCCoCfg         As ADODB.Recordset
Private porstCOCta          As ADODB.Recordset
Private nFormato As Integer
Private psConnStrgSele As String
Public Rstfiltro As ADODB.Recordset
Dim Valx As Boolean

']
Dim ApExcel As Variant
Dim AHojas() As String
Dim ANombres() As String
Dim ACostos() As String
Dim NCostos() As String
Dim strsql As String


Private Sub chkRango_Click()
  fraRngPeriodo.Enabled = (chkRango.Value = vbChecked)
End Sub


Private Sub Form_Load()
  On Error GoTo Err
  Valx = False
  
  If gsNvlUsr <> 0 Then
    chkpresupuesto.Visible = False
    cmdexcel.Visible = False
  End If
  
  Dim dnContador As Integer

  Set cnn = New ADODB.Connection
  If ps_Puerto = "" Then
     cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";connection="
  Else
     cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";Port=" & ps_Puerto & ";connection="
  End If
  cnn.CursorLocation = adUseClient
  cnn.Open

  '[Recordsets.  'Cambiar.
  Set pocnnMain = New ADODB.Connection
  Set porstMRp = New ADODB.Recordset
  Set porstCOCta = New ADODB.Recordset
  Set porstCCoCfg = New ADODB.Recordset
  Set porstMRpRs = New ADODB.Recordset
   
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
    
  With porstMRp
    .ActiveConnection = pocnnMain
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
    
  With porstCOCta
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
    .Source = .Source & "FROM CoCta "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodCta"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  
  psConnStrgSele = "SELECT DISTINCTROW codcfg, detcfg "
  psConnStrgSele = psConnStrgSele & "FROM coccocfg "
  psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
  psConnStrgSele = psConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
  With porstCCoCfg
    .ActiveConnection = pocnnMain
    .Source = psConnStrgSele
    .Source = .Source & "AND tipofmt=" & nFormato & " "
    .Source = .Source & "ORDER BY codcfg"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  txtDato(2).DataField = "codcfg"
  txtDato(2).MaxLength = porstCCoCfg.Fields(txtDato(2).DataField).DefinedSize
  
  With porstMRpRs
    .ActiveConnection = pocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockBatchOptimistic
    .Source = "SELECT * "
    .Source = .Source & "FROM cotmprpt "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "'"
    '        .Open
  End With

 ']

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :", "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :", "Beginning :", "End :")
  Next nElemento
  
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  optFormato(0).Caption = Choose(gsIdioma, "Cuentas :", "Accounts :")
  optFormato(1).Caption = Choose(gsIdioma, "Cen. Costos :", "Cost Center :")
  fraFormato.Caption = Choose(gsIdioma, "Formato", "Format")
  fraAlcance.Caption = Choose(gsIdioma, "Alcance", "Scope")
  optAlcance(0).Caption = Choose(gsIdioma, "al mes", "to month")
  optAlcance(1).Caption = Choose(gsIdioma, "del mes", "from month")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  cmdCCostos.Caption = Choose(gsIdioma, "C.Costos", "Cost Center")
  fraRngPeriodo.Caption = Choose(gsIdioma, "Rango Periodos", "Range of Periods")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND
    
 
  '[Datos predeterminados.              'Cambiar.
  optFormato(0).Value = True
  ' Configuro los controles de año y mes
  For dnContador = (Val(gsAnoAct) - 9) To Val(gsAnoAct)
    cmbPeriodo(0).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador
    cmbPeriodo(1).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador
  Next dnContador
  cmbPeriodo(0).ListIndex = 9
  cmbPeriodo(1).ListIndex = 9
  
  For dnContador = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      cmbPeriodo(2).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
      cmbPeriodo(3).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
    Else
      cmbPeriodo(2).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
      cmbPeriodo(3).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
    End If
  Next dnContador
  cmbPeriodo(2).ListIndex = Val(gsMesAct)
  cmbPeriodo(3).ListIndex = Val(gsMesAct)
  fraRngPeriodo.Enabled = False
 ']
  
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
   porstCCoCfg.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstCCoCfg = Nothing
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
  Dim dnContador As Integer, nNivel As Integer
  Dim n_Index As Integer, nRegistro As Integer
  Dim CadCrystal As String, s_Moneda As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_MesIni As Integer, n_MesFin As Integer
  Dim porsClone As ADODB.Recordset
  
  Dim Rstfiltro As ADODB.Recordset, Rstdatos As ADODB.Recordset
  Dim contador1 As Integer, contador2 As Integer
  Dim sql As String
  
  Set Rstfiltro = New ADODB.Recordset
  Set Rstdatos = New ADODB.Recordset
      
  ' Verifico los datos ingresados
  If txtDato(2).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(2).SetFocus: Exit Sub
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  
  ppHabilitacion False
   
  Set porsClone = New ADODB.Recordset
  With porsClone
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    '.CursorLocation = adUseClient   'Es el Default.
    .Source = "SELECT NumOrd, CodCCo, DetCCo, Nivel "
    .Source = .Source & "FROM CoCCoCfg "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND TipoFmt='" & nFormato & "' "
    .Source = .Source & "AND codcfg='" & txtDato(2).Text & "' "
    .Source = .Source & "ORDER BY NumOrd"
    .Open
  End With

  ' Elimino y genero el archivo del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCco", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCco') DROP TABLE #trpRngEfiCco")
  'CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE trpRngEfiCco (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCco (")
  CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE trpRngEfiCco (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCco (")
  CadCrystal = CadCrystal & "codcta varchar(16) Not Null,"
  CadCrystal = CadCrystal & "detcta varchar(60) Default Null,"
  CadCrystal = CadCrystal & "x00 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x01 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x02 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x03 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x04 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x05 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x06 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x07 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x08 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x09 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x10 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x11 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x12 decimal(12,2) Not Null Default '0.00'," '
  CadCrystal = CadCrystal & "xXX decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "yYY decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "xTotal decimal(12,2) Not Null Default '0.00') "
  pocnnMain.Execute CadCrystal
  
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCcox", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCco') DROP TABLE #trpRngEfiCcox")
  'CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE trpRngEfiCcox (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCcox (")
  CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE trpRngEfiCcox (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCcox (")
  CadCrystal = CadCrystal & "codcta varchar(16) Not Null,"
  CadCrystal = CadCrystal & "detcta varchar(60) Default Null,"
  CadCrystal = CadCrystal & "x00 decimal(12,2) Default Null,"
  CadCrystal = CadCrystal & "x01 decimal(12,2) Default Null,"
  CadCrystal = CadCrystal & "x02 decimal(12,2) Default Null,"
  CadCrystal = CadCrystal & "x03 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x04 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x05 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x06 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x07 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x08 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x09 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x10 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x11 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x12 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "xXX decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "yYY decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "xTotal decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "aaa decimal(12,2) Not Null Default '0.00', "
  CadCrystal = CadCrystal & "detagr varchar(60) Not Null, "
  CadCrystal = CadCrystal & "bbb decimal(12,2) Default '0.00') "
  pocnnMain.Execute CadCrystal
   
  ' Obtengo el nivel de analisis si existe registros
  If porsClone.RecordCount > 0 Then
    nRegistro = porsClone.RecordCount
    nNivel = porsClone!nivel
    
    For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
      s_Ano = Trim$(dnContador)
      n_MesIni = Val(IIf(optAlcance(0).Value, 0, gsMesAct))
      n_MesFin = Val(gsMesAct)
      If chkRango.Value = vbChecked Then
        n_MesIni = Val(IIf(s_Ano = s_AnoIni, cmbPeriodo(2).ListIndex, 1))
        n_MesFin = Val(IIf(s_Ano = s_AnoFin, cmbPeriodo(3).ListIndex, 12))
      End If
      ' Acumulación de saldos
      s_SaldoDeb = "": s_SaldoHab = ""
      For n_Index = n_MesIni To n_MesFin
        s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
        s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      Next n_Index
    
      ' Inserto los registros
      CadCrystal = "INSERT INTO " & ps_Prefijo & "trpRngEfiCco "
      If nFormato = "0" Then
        CadCrystal = CadCrystal & "SELECT a.codcta, MAX(" & Choose(gsIdioma, "b.detcta", "b.detctax") & ") AS detcta, "
      Else
        CadCrystal = CadCrystal & "SELECT a.codcco AS codcta, MAX(" & Choose(gsIdioma, "b.DetCCo", "b.DetCCox") & ") AS detcta, "
      End If
      ' primer paso: Acumulo las columnas por centro de costos o cuenta
      porsClone.MoveFirst
      Do While Not porsClone.EOF
        CadCrystal = CadCrystal & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
        CadCrystal = CadCrystal & "CASE a." & Choose(nFormato + 1, "codcco", "codcta") & " WHEN '" & porsClone!codcco & "' THEN "
        CadCrystal = CadCrystal & "(" & s_SaldoDeb & ")-(" & s_SaldoHab & ") "
        CadCrystal = CadCrystal & "ELSE 0 END, 0)), 2) AS " & Trim("x" & porsClone!numord) & ", "
        porsClone.MoveNext
      Loop
    
      ' segundo paso: Registro otros ultimo registro
      nRegistro = nRegistro - IIf(nRegistro > 1, 1, 0)
      For n_Index = nRegistro To 12
        CadCrystal = CadCrystal & "0 AS " & Trim("x" & Format(n_Index, "00")) & ", "
      Next n_Index
      nRegistro = nRegistro + IIf(nRegistro > 1, 1, 0)
      
      'Paso Adicional : Septiembre del 2008 : ma 08-07-2011
      CadCrystal = CadCrystal & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
      CadCrystal = CadCrystal & "CASE WHEN x.indpdocpr=1 THEN "
      CadCrystal = CadCrystal & "(CASE WHEN a." & Choose(nFormato + 1, "codcco", "codcta") & " IN (SELECT DISTINCT cfg.codcco FROM coccocfg cfg "
      CadCrystal = CadCrystal & "WHERE cfg.codemp=a.codemp AND cfg.pdoano=a.pdoano "
      CadCrystal = CadCrystal & "AND cfg.tipofmt='" & nFormato & "' AND cfg.codcfg='" & txtDato(2).Text & "') "
      CadCrystal = CadCrystal & "THEN (" & s_SaldoDeb & ")-(" & s_SaldoHab & ") ELSE 0 END) "
      CadCrystal = CadCrystal & "ELSE 0 END, 0)), 2) AS yYY, "
        
      ' tercer paso: genero le total general
      CadCrystal = CadCrystal & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
      CadCrystal = CadCrystal & "(" & s_SaldoDeb & ")-(" & s_SaldoHab & "), 0)), 2) AS xTotal "
      ' cuarto paso: seleccion de tablas y condicion
      CadCrystal = CadCrystal & "FROM (coccoacu a "
      CadCrystal = CadCrystal & "INNER JOIN cocco x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.codcco=x.codcco "
      
      If nFormato = "0" Then
        CadCrystal = CadCrystal & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.codcta=b.codcta) "
        CadCrystal = CadCrystal & "WHERE a.codemp='" & gsCodEmp & "' "
        CadCrystal = CadCrystal & "AND a.pdoano='" & s_Ano & "' "
        CadCrystal = CadCrystal & "AND (a.codcta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "') "
        CadCrystal = CadCrystal & "AND b.tpocta='" & TPOCTA_TRA & "' "
        CadCrystal = CadCrystal & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.codcco)=" & nNivel & " "
        CadCrystal = CadCrystal & "GROUP BY a.codcta "
        CadCrystal = CadCrystal & "ORDER BY a.codcta"
      Else
        CadCrystal = CadCrystal & "LEFT JOIN cocco b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCCo=b.CodCCo) "
        CadCrystal = CadCrystal & "WHERE a.codemp='" & gsCodEmp & "' "
        CadCrystal = CadCrystal & "AND a.pdoano='" & s_Ano & "' "
        CadCrystal = CadCrystal & "AND (a.codcco BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "') "
        CadCrystal = CadCrystal & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(b.CodCCo)=" & Right(gsNivCCo, 1) & " "
        CadCrystal = CadCrystal & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodCta)=" & nNivel & ""
        CadCrystal = CadCrystal & "GROUP BY a.codcco "
        CadCrystal = CadCrystal & "ORDER BY a.codcco"
      End If
      pocnnMain.Execute CadCrystal
    Next dnContador
  End If
  
  ' Obtengo los registros del reporte
  CadCrystal = "INSERT INTO " & ps_Prefijo & "trpRngEfiCcox "
  CadCrystal = CadCrystal & "SELECT codcta, detcta, "
  CadCrystal = CadCrystal & "ROUND(SUM(x00), 2) AS x00, "
  CadCrystal = CadCrystal & "ROUND(SUM(x01), 2) AS x01, "
  CadCrystal = CadCrystal & "ROUND(SUM(x02), 2) AS x02, "
  CadCrystal = CadCrystal & "ROUND(SUM(x03), 2) AS x03, "
  CadCrystal = CadCrystal & "ROUND(SUM(x04), 2) AS x04, "
  CadCrystal = CadCrystal & "ROUND(SUM(x05), 2) AS x05, "
  CadCrystal = CadCrystal & "ROUND(SUM(x06), 2) AS x06, "
  CadCrystal = CadCrystal & "ROUND(SUM(x07), 2) AS x07, "
  CadCrystal = CadCrystal & "ROUND(SUM(x08), 2) AS x08, "
  CadCrystal = CadCrystal & "ROUND(SUM(x09), 2) AS x09, "
  CadCrystal = CadCrystal & "ROUND(SUM(x10), 2) AS x10, "
  CadCrystal = CadCrystal & "ROUND(SUM(x11), 2) AS x11, "
  CadCrystal = CadCrystal & "ROUND(SUM(x12), 2) AS x12, "
  CadCrystal = CadCrystal & "ROUND(SUM(xXX), 2) AS xXX, "
  CadCrystal = CadCrystal & "ROUND(SUM(yYY), 2) AS yYY, "
  CadCrystal = CadCrystal & "ROUND(SUM(xTotal), 2) AS xTotal, "
  CadCrystal = CadCrystal & "ROUND(SUM(0), 2) AS aaa, "
  CadCrystal = CadCrystal & "'' AS detagr, "
  CadCrystal = CadCrystal & "0  AS bbb "
  CadCrystal = CadCrystal & "FROM " & ps_Prefijo & "trpRngEfiCco "
  CadCrystal = CadCrystal & "GROUP BY codcta, detcta "
  CadCrystal = CadCrystal & "ORDER BY codcta "
  pocnnMain.Execute CadCrystal
  
  sql = " select codcta from " & ps_Prefijo & "trpRngEfiCcox"
  Rstfiltro.Open sql, pocnnMain, adOpenStatic, adLockOptimistic

  If Rstfiltro.RecordCount = 0 Then
     Exit Sub
  Else
     Rstfiltro.MoveFirst
      'recorre todo el recordset
     For contador1 = 0 To Rstfiltro.RecordCount - 1
      sql = "SELECT ROUND(SUM(impcta_" & s_Moneda & "), 2) FROM copdocprcta WHERE copdocprcta.codemp='" & gsCodEmp & "' "
      sql = sql & "AND copdocprcta." & Choose(nFormato + 1, "codcco", "codcta") & " IN (SELECT DISTINCT CoCCoCfg.codcco FROM CoCCoCfg WHERE CoCCoCfg.codemp='" & gsCodEmp & "' "
      sql = sql & "AND CoCCoCfg.pdoano='" & gsAnoAct & "' "
      sql = sql & "AND CoCCoCfg.TipoFmt='" & nFormato & "' "
      sql = sql & "AND CoCCoCfg.codcfg='" & txtDato(2).Text & "') "
      sql = sql & "AND copdocprcta.codcta='" & Rstfiltro.Fields(0).Value & "' "
      If chkRango.Value = vbChecked Then
        sql = sql & "AND concat(copdocprcta.pdoano,copdocprcta.mespvs)>= '" & s_AnoIni & s_Mes & "' AND concat(copdocprcta.pdoano,copdocprcta.mespvs) <= '" & s_AnoFin & Format(cmbPeriodo(3).ListIndex, "00") & "' "
        sql = sql & "GROUP BY copdocprcta.codemp, copdocprcta.codcta, left(copdocprcta.codcco,2)"
      Else
        sql = sql & "AND copdocprcta.pdoano='" & gsAnoAct & "' AND copdocprcta.mespvs<='" & gsMesAct & " ' "
        sql = sql & "GROUP BY copdocprcta.codemp, copdocprcta.pdoano, copdocprcta.codcta, left(copdocprcta.codcco,2)"
      End If
      Rstdatos.Open sql, cnn, adOpenStatic, adLockOptimistic
      If Rstdatos.RecordCount = 0 Then
      Else
        Rstdatos.MoveFirst
        For contador2 = 0 To Rstdatos.RecordCount - 1
          With porstMRp
            If .State = adStateOpen Then .Close
            CadCrystal = "Update " & ps_Prefijo & "trpRngEfiCcox set " & ps_Prefijo & "trpRngEfiCcox.aaa = " & Rstdatos.Fields(0).Value & " where codcta='" & Rstfiltro.Fields(0).Value & "'"
            .Source = CadCrystal
            .Open
          End With
          Rstdatos.MoveNext
        Next contador2
      End If
      Rstdatos.Close
      Rstfiltro.MoveNext
    Next contador1
  End If
  Rstfiltro.Close
  
  Dim sentencia As String
  Dim sumatoria As String
  Dim desde As Integer
 
  If chkpresupuesto.Value = Checked Then
    sentencia = "SUM("
    sumatoria = ""
    'Momentaneo
    If Right(cmbPeriodo(1).Text, 4) > Right(cmbPeriodo(0).Text, 4) Then
      For desde = 1 To 12
        sentencia = sentencia & "pre.imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(12), "", "+")
        sumatoria = sumatoria & "pre.imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(12), "", "+")
      Next
      sentencia = sentencia & ")"
    Else
      For desde = 1 To Int(gsMesAct)
        sentencia = sentencia & "pre.imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(gsMesAct), "", "+")
        sumatoria = sumatoria & "pre.imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(gsMesAct), "", "+")
      Next
      sentencia = sentencia & ")"
    End If
    
    If gsMesAct <> "00" Then
      With porstMRp
        If .State = adStateOpen Then .Close
        CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x SET x.bbb=("
        If optAlcance(0).Value = True Then
          'CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.bbb=(select " & sentencia & " FROM copsp where codcta=x.codcta and left(codcco,2)='" & txtDato(2) & "' and pdoano='" & s_Ano & "' and codemp='" & gsCodEmp & "' )  "
          CadCrystal = CadCrystal & "SELECT ROUND(" & sentencia & ", 2) FROM copsp pre "
        Else
          'CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.bbb=(select imp" & s_Moneda & "_" & gsMesAct & " FROM copsp where codcta=x.codcta and left(codcco,2)='" & txtDato(2) & "' and pdoano='" & s_Ano & "' and codemp='" & gsCodEmp & "' )  "
          CadCrystal = CadCrystal & "SELECT ROUND(pre.imp" & s_Moneda & "_" & gsMesAct & ", 2) FROM copsp pre "
        End If
        CadCrystal = CadCrystal & "WHERE pre.codemp='" & gsCodEmp & "' "
        CadCrystal = CadCrystal & "AND pre.pdoano>='" & Right(cmbPeriodo(0).Text, 4) & "' AND pre.pdoano<='" & Right(cmbPeriodo(1).Text, 4) & "' "
        CadCrystal = CadCrystal & "AND pre.codcta=x.codcta "
        CadCrystal = CadCrystal & "AND left(pre.codcco, 2) IN (SELECT DISTINCT LEFT(cfg.codcco, 2) FROM coccocfg cfg "
        CadCrystal = CadCrystal & "WHERE cfg.codemp=pre.codemp AND cfg.pdoano=pre.pdoano "
        CadCrystal = CadCrystal & "AND cfg.tipofmt='" & nFormato & "' AND cfg.codcfg='" & txtDato(2).Text & "'))"
        .Source = CadCrystal
        .Open
      End With
      ' 12-11-2011 Inserto cuentas sin movimiento
      With porstMRp
        If .State = adStateOpen Then .Close
        CadCrystal = "INSERT INTO trpRngEfiCcox SELECT pre.codcta, b.detcta, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,'', "
        If optAlcance(0).Value = True Then
          CadCrystal = CadCrystal & "ROUND(" & sumatoria & ", 2) "
          CadCrystal = CadCrystal & "FROM copsp pre INNER JOIN cocta b ON b.codemp=pre.codemp AND b.pdoano=pre.pdoano AND b.codcta=pre.codcta "
          CadCrystal = CadCrystal & "WHERE pre.codemp='" & gsCodEmp & "' "
          CadCrystal = CadCrystal & "AND pre.pdoano>='" & Right(cmbPeriodo(0).Text, 4) & "' AND pre.pdoano<='" & Right(cmbPeriodo(1).Text, 4) & "' "
          CadCrystal = CadCrystal & "AND pre.codcta not in (select x.codcta from trpRngEfiCcox x) "
          CadCrystal = CadCrystal & "AND LEFT(pre.codcco, 2) IN (SELECT DISTINCT LEFT(cfg.codcco, 2) FROM coccocfg cfg "
          CadCrystal = CadCrystal & "WHERE cfg.codemp=pre.codemp AND cfg.pdoano=pre.pdoano "
          CadCrystal = CadCrystal & "AND cfg.tipofmt='" & nFormato & "' AND cfg.codcfg='" & txtDato(2).Text & "')"
        Else
          CadCrystal = CadCrystal & "ROUND(pre.imp" & s_Moneda & "_" & gsMesAct & ", 2) "
          CadCrystal = CadCrystal & "FROM copsp pre INNER JOIN cocta b ON b.codemp=pre.codemp AND b.pdoano=pre.pdoano AND b.codcta=pre.codcta "
          CadCrystal = CadCrystal & "WHERE pre.codemp='" & gsCodEmp & "' "
          CadCrystal = CadCrystal & "AND pre.pdoano='" & s_Ano & "' "
          CadCrystal = CadCrystal & "AND pre.codcta not in (select x.codcta from trpRngEfiCcox x) "
          CadCrystal = CadCrystal & "AND LEFT(pre.codcco, 2) IN (SELECT DISTINCT LEFT(cfg.codcco, 2) FROM coccocfg cfg "
          CadCrystal = CadCrystal & "WHERE cfg.codemp=pre.codemp AND cfg.pdoano=pre.pdoano "
          CadCrystal = CadCrystal & "AND cfg.tipofmt='" & nFormato & "' AND cfg.codcfg='" & txtDato(2).Text & "')"
        End If
        .Source = CadCrystal
        .Open
      End With
    End If
    
 'Else
 'For desde = cmbPeriodo(0).Text To cmbPeriodo(1).Text
 'Next
 'End If
  End If
 
  With porstMRp
    If .State = adStateOpen Then .Close
    CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.detagr=(select detcta FROM cocta where codcta=left(x.codcta,2) and pdoano='" & s_Ano & "' and codemp='" & gsCodEmp & "' ) "
    .Source = CadCrystal
    .Open
  End With
     
  ' Obtengo los registros del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    CadCrystal = "SELECT codcta, detcta, "
    CadCrystal = CadCrystal & "ROUND(SUM(x00), 2) AS x00, "
    CadCrystal = CadCrystal & "ROUND(SUM(x01), 2) AS x01, "
    CadCrystal = CadCrystal & "ROUND(SUM(x02), 2) AS x02, "
    CadCrystal = CadCrystal & "ROUND(SUM(x03), 2) AS x03, "
    CadCrystal = CadCrystal & "ROUND(SUM(x04), 2) AS x04, "
    CadCrystal = CadCrystal & "ROUND(SUM(x05), 2) AS x05, "
    CadCrystal = CadCrystal & "ROUND(SUM(x06), 2) AS x06, "
    CadCrystal = CadCrystal & "ROUND(SUM(x07), 2) AS x07, "
    CadCrystal = CadCrystal & "ROUND(SUM(x08), 2) AS x08, "
    CadCrystal = CadCrystal & "ROUND(SUM(x09), 2) AS x09, "
    CadCrystal = CadCrystal & "ROUND(SUM(x10), 2) AS x10, "
    CadCrystal = CadCrystal & "ROUND(SUM(x11), 2) AS x11, "
    CadCrystal = CadCrystal & "ROUND(SUM(x12), 2) AS x12, "
    CadCrystal = CadCrystal & "ROUND(SUM(xXX), 2) AS xXX, "
    CadCrystal = CadCrystal & "ROUND(SUM(yYY), 2) AS yYY, "
    CadCrystal = CadCrystal & "ROUND(SUM(xTotal), 2) AS xTotal, "
    CadCrystal = CadCrystal & "ROUND(SUM(aaa), 2) AS aaa, "
    CadCrystal = CadCrystal & "detagr AS detagr, "
    CadCrystal = CadCrystal & "if(bbb is null , 0 ,bbb) AS bbb "
    CadCrystal = CadCrystal & "FROM " & ps_Prefijo & "trpRngEfiCcox "
    CadCrystal = CadCrystal & "GROUP BY codcta, detcta "
    If Valx = False Then
        CadCrystal = CadCrystal & "HAVING (x00+x01+x02+x03+x04+x05+x06+x07+x08+x09+x10+bbb) <> 0.00 OR ROUND(SUM(aaa), 2) <> 0 "
    Else
        CadCrystal = CadCrystal & "HAVING (x00+x01+x02+x03+x04+x05+x06+x07+x08+x09+x10+bbb) >= 0.00 OR ROUND(SUM(aaa), 2) >= 0 "
    End If
    CadCrystal = CadCrystal & "ORDER BY codcta"
    .Source = CadCrystal
    .Open
  End With
  
  CadCrystal = IIf(chkRango.Value = vbChecked, cmbPeriodo(2).Text & " - " & cmbPeriodo(0).Text, "")
  'usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_GRAF, PRN_DEST_MATR)
  
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRptPresup frmMain.rptMain, Me.Caption & " -" & Trim(lblDatoDeta(2).Caption) & "- (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp, IIf(chkpresupuesto.Value = Checked, 0, 1)
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
       If chkDivisoria.Value = vbChecked Then
      .ReportFileName = gsRutRpt & "rptREFiCCoD.rpt"
      Else
      .ReportFileName = gsRutRpt & "rptREFiCCo.rpt"
      End If
      .Formulas(5) = "mPeriodo='" & CadCrystal & " " & IIf(optAlcance(0).Value, Choose(gsIdioma, "Acumulado - ", "Accrued - "), "") & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
      If porsClone.RecordCount > 0 Then
        porsClone.MoveFirst
        Do While Not porsClone.EOF
          Select Case Trim("" & porsClone!numord)
           Case "00": .Formulas(6) = "c1='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(7) = "p1='" & Trim("" & porsClone!codcco) & "'"
           Case "01": .Formulas(8) = "c2='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(9) = "p2='" & Trim("" & porsClone!codcco) & "'"
           Case "02": .Formulas(10) = "c3='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(11) = "p3='" & Trim("" & porsClone!codcco) & "'"
           Case "03": .Formulas(12) = "c4='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(13) = "p4='" & Trim("" & porsClone!codcco) & "'"
           Case "04": .Formulas(14) = "c5='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(15) = "p5='" & Trim("" & porsClone!codcco) & "'"
           Case "05": .Formulas(16) = "c6='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(17) = "p6='" & Trim("" & porsClone!codcco) & "'"
           Case "06": .Formulas(18) = "c7='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(19) = "p7='" & Trim("" & porsClone!codcco) & "'"
           Case "07": .Formulas(20) = "c8='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(21) = "p8='" & Trim("" & porsClone!codcco) & "'"
           Case "08": .Formulas(22) = "c9='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(23) = "p9='" & Trim("" & porsClone!codcco) & "'"
           Case "09": .Formulas(24) = "c10='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(25) = "p10='" & Trim("" & porsClone!codcco) & "'"
           Case "10": .Formulas(26) = "c11='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(27) = "p11='" & Trim("" & porsClone!codcco) & "'"
           Case "XX": .Formulas(28) = "c12='" & Trim("" & porsClone!detcco) & "'"
            .Formulas(29) = "p12='" & Trim("" & porsClone!codcco) & "'"
          End Select
          porsClone.MoveNext
        Loop
        .Formulas(41) = "xTotal={trpteficco.xTOTAL}"
      End If
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRpRs
      .LoadReport gsRutRpt & "rptREFiCCo.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = IIf(optAlcance(0).Value = True, "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct, Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct)
      If porsClone.RecordCount > 0 Then
        porsClone.MoveFirst
        Do While Not porsClone.EOF
          Select Case Trim("" & porsClone!numord)
            Case "00": .Parameters("pc1") = Trim("" & porsClone!detcco): .Parameters("pt1") = Trim("" & porsClone!codcco)
            Case "01": .Parameters("pc2") = Trim("" & porsClone!detcco): .Parameters("pt2") = Trim("" & porsClone!codcco)
            Case "02": .Parameters("pc3") = Trim("" & porsClone!detcco): .Parameters("pt3") = Trim("" & porsClone!codcco)
            Case "03": .Parameters("pc4") = Trim("" & porsClone!detcco): .Parameters("pt4") = Trim("" & porsClone!codcco)
            Case "04": .Parameters("pc5") = Trim("" & porsClone!detcco): .Parameters("pt5") = Trim("" & porsClone!codcco)
            Case "05": .Parameters("pc6") = Trim("" & porsClone!detcco): .Parameters("pt6") = Trim("" & porsClone!codcco)
            Case "06": .Parameters("pc7") = Trim("" & porsClone!detcco): .Parameters("pt7") = Trim("" & porsClone!codcco)
            Case "07": .Parameters("pc8") = Trim("" & porsClone!detcco): .Parameters("pt8") = Trim("" & porsClone!codcco)
            Case "08": .Parameters("pc9") = Trim("" & porsClone!detcco): .Parameters("pt9") = Trim("" & porsClone!codcco)
            Case "09": .Parameters("pc10") = Trim("" & porsClone!detcco): .Parameters("pt10") = Trim("" & porsClone!codcco)
            Case "10": .Parameters("pc11") = Trim("" & porsClone!detcco): .Parameters("pt11") = Trim("" & porsClone!codcco)
            Case "XX": .Parameters("pc12") = Trim("" & porsClone!detcco): .Parameters("pt12") = Trim("" & porsClone!codcco)
          End Select
          porsClone.MoveNext
        Loop
      End If
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
  
  
  porsClone.Close
  Set porsClone = Nothing
  ' Elimino y genero el archivo del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCco", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCco') DROP TABLE #trpRngEfiCco")
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCcox", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCcox') DROP TABLE #trpRngEfiCcox")
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

Private Sub optFormato_Click(Index As Integer)
  Dim nIndex As Integer
  
  nFormato = Index
  With porstCOCta
    If .State = adStateOpen Then .Close
    If nFormato = 0 Then
      .Source = "SELECT codcta, " & Choose(gsIdioma, "detcta", "detctax") & " AS detcta "
      .Source = .Source & "FROM cocta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY codcta"
    Else
      .Source = "SELECT codcco AS codcta, " & Choose(gsIdioma, "detcco", "detccox") & " AS detcta "
      .Source = .Source & "FROM cocco "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY codcco"
    End If
    .Open
  End With
  
  ' Inicializo y cambio las etiquetas
  For nIndex = 0 To 1
    txtDato(nIndex).DataField = "codcta"
    txtDato(nIndex).MaxLength = porstCOCta.Fields(txtDato(nIndex).DataField).DefinedSize
  Next nIndex
  cmdCCostos.Caption = Choose(nFormato + 1, Choose(gsIdioma, "C.Costos", "Cost Center"), Choose(gsIdioma, "Cuentas", "Accounts"))
  
  ' Infroamcion de formatos
  With porstCCoCfg
    If .State = adStateOpen Then .Close
    .Source = psConnStrgSele & "AND tipofmt=" & nFormato & " "
    .Source = .Source & "ORDER BY codcfg"
    .Open
  End With
  txtDato(2).Text = ""
  lblDatoDeta(2).Caption = ""
  
  ' Límites de rangos.
  porstCOCta.MoveLast
  txtDato(1).Text = porstCOCta(txtDato(1).DataField)
  porstCOCta.MoveFirst
  txtDato(0).Text = porstCOCta(txtDato(0).DataField)
  
  ' Busca detalle de códigos
  If txtDato(0).Text <> "" Then ppAyuDet 0
  If txtDato(1).Text <> "" Then ppAyuDet 1

End Sub

Private Sub Todos_Click()
    If Todos.Value = Checked Then
        Valx = True
    Else
        Valx = False
    End If
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

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  'Select Case Index    'Completa con ceros a la izquierda.
  'Case 0, 1                           'Cambiar (añadir índices).
  '   If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
  '      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
  '   End If
  'End Select

  Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2           'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
    If nFormato = 0 Then
      modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
    Else
      modAyuBus.CCo_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
    End If
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2                           'Cambiar (añadir índices).
    modAyuBus.Cfg_Cod "tipofmt=" & nFormato & "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraFormato.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraFormato.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 0, 1
    If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
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
   Case 2
    If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
    With porstCCoCfg
      .MoveFirst
      .Find "codcfg='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcfg), "", !detcfg)
      End If
    End With
  End Select
End Function

'[Propio del formulario.
Private Sub cmdCCostos_Click()
  frmMEFiCCoGrd.Show vbModal
  porstCCoCfg.Requery
End Sub
']

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdexcel.Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

  'Controles del formulario.
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property
Private Sub cmdexcel_Click()

Dim i As Integer
Dim sql As String
Dim rshojas As New ADODB.Recordset
Dim s_Filtro_Annos As String

ReDim AHojas(20)
ReDim ANombres(20)

Set ApExcel = CreateObject("Excel.application")
ApExcel.Visible = False

s_Filtro_Annos = ""
For i = Right(cmbPeriodo(0), 4) To Right(cmbPeriodo(1), 4)
    s_Filtro_Annos = s_Filtro_Annos & i & IIf(i = Right(cmbPeriodo(1), 4), "", ",")
Next

'ULTIMO AÑO
s_Filtro_Annos = Right(cmbPeriodo(1), 4)
s_Filtro_Annos = " pdoano in ( " & s_Filtro_Annos & " )"

ApExcel.Workbooks.Add
    
sql = "select distinctrow codcfg, detcfg from coccocfg where codemp='" & gsCodEmp & "' and " & s_Filtro_Annos & " and tipofmt=0 order by 1 asc "
rshojas.Open sql, cnn, adOpenDynamic, adLockPessimistic
i = 1
Do Until rshojas.EOF
    
    If i < 4 Then
        ApExcel.Sheets("Hoja" & i).Name = rshojas(0).Value
        AHojas(i - 1) = rshojas(0).Value
        ANombres(i - 1) = rshojas(1).Value
    Else
        ApExcel.Sheets.Add
        ApExcel.Sheets("Hoja" & i).Name = rshojas(0).Value
        AHojas(i - 1) = rshojas(0).Value
        ANombres(i - 1) = rshojas(1).Value
    End If
    
    i = i + 1
    rshojas.MoveNext
    
Loop

ProcesarHojas

MsgBox ("Proceso de Exportacion a Excel, terminado")
ApExcel.Visible = True
error:

End Sub
Sub ProcesarHojas()
Dim k As Integer
For k = LBound(AHojas) To UBound(AHojas)
    If AHojas(k) <> "" Then
        
    ApExcel.ActiveWindow.Zoom = 65
    ApExcel.Sheets(AHojas(k)).Select
    ApExcel.Range("A1").Select
    ApExcel.ActiveCell = gsRazEmp & "       Acumulado - Desde " & cmbPeriodo(2) & " del " & Right(cmbPeriodo(0), 4) & " Hasta " & cmbPeriodo(3) & " del " & Right(cmbPeriodo(1), 4) & ""
    With ApExcel.Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        '.Underline = xlUnderlineStyleNone
    End With
    ApExcel.Selection.Font.Bold = True

    ApExcel.Range("A2").Select
    ApExcel.ActiveCell = cboTpoMon.Text
    With ApExcel.Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        '.Underline = xlUnderlineStyleNone
    End With
  
    ApExcel.Range("A3").Select
    ApExcel.ActiveCell = ANombres(k)
    With ApExcel.Selection.Font
        .Name = "Arial"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        '.Underline = xlUnderlineStyleNone
    End With
    ApExcel.Selection.Font.Bold = True

    procesardatos (AHojas(k))
    
    darformato (AHojas(k))

    End If
Next
End Sub
Sub procesardatos(hoja As String)

Dim i As Integer
Dim j As Integer
Dim cols As Integer

Dim dnContador As Integer, nNivel As Integer
  Dim n_Index As Integer, nRegistro As Integer
  Dim CadCrystal As String, s_Moneda As String
  Dim porsClone As ADODB.Recordset
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_MesIni As Integer, n_MesFin As Integer
  
  Dim Rstfiltro As ADODB.Recordset
  Dim Rstdatos As ADODB.Recordset
  Dim contador1 As Integer
  Dim contador2 As Integer
  Dim sql As String
  Set Rstfiltro = New ADODB.Recordset
  Set Rstdatos = New ADODB.Recordset
      
  ' Verifico los datos ingresados
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  
  ppHabilitacion False
   
  Set porsClone = New ADODB.Recordset
  With porsClone
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    '.CursorLocation = adUseClient   'Es el Default.
    .Source = "SELECT NumOrd, CodCCo, DetCCo, Nivel "
    .Source = .Source & "FROM CoCCoCfg "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND TipoFmt='" & nFormato & "' "
    .Source = .Source & "AND codcfg='" & hoja & "' "
    .Source = .Source & "ORDER BY NumOrd"
    .Open
  End With

  ' Elimino y genero el archivo del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCco", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCco') DROP TABLE #trpRngEfiCco")
  'CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE trpRngEfiCco (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCco (")
  CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE trpRngEfiCco (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCco (")
  CadCrystal = CadCrystal & "codcta varchar(16) Not Null,"
  CadCrystal = CadCrystal & "detcta varchar(60) Default Null,"
  CadCrystal = CadCrystal & "x00 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x01 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x02 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x03 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x04 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x05 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x06 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x07 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x08 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x09 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x10 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x11 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x12 decimal(12,2) Not Null Default '0.00'," '
  CadCrystal = CadCrystal & "xXX decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "yYY decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "xTotal decimal(12,2) Not Null Default '0.00') "
  pocnnMain.Execute CadCrystal
  
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCcox", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCco') DROP TABLE #trpRngEfiCcox")
  'CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE trpRngEfiCcox (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCcox (")
  CadCrystal = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE trpRngEfiCcox (", "CREATE TABLE " & ps_Prefijo & "trpRngEfiCcox (")
  CadCrystal = CadCrystal & "codcta varchar(16) Not Null,"
  CadCrystal = CadCrystal & "detcta varchar(60) Default Null,"
  CadCrystal = CadCrystal & "x00 decimal(12,2) Default Null,"
  CadCrystal = CadCrystal & "x01 decimal(12,2) Default Null,"
  CadCrystal = CadCrystal & "x02 decimal(12,2) Default Null,"
  CadCrystal = CadCrystal & "x03 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x04 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x05 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x06 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x07 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x08 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x09 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x10 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x11 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "x12 decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "xXX decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "yYY decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "xTotal decimal(12,2) Not Null Default '0.00',"
  CadCrystal = CadCrystal & "aaa decimal(12,2) Not Null Default '0.00', "
  CadCrystal = CadCrystal & "detagr varchar(60) Not Null, "
  CadCrystal = CadCrystal & "bbb decimal(12,2) Default '0.00') "
  pocnnMain.Execute CadCrystal
   
  ' Obtengo el nivel de analisis si existe registros
  If porsClone.RecordCount > 0 Then
    
    nRegistro = porsClone.RecordCount
    nNivel = porsClone!nivel
    
    For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
      s_Ano = Trim$(dnContador)
      n_MesIni = Val(IIf(optAlcance(0).Value, 0, gsMesAct))
      n_MesFin = Val(gsMesAct)
      If chkRango.Value = vbChecked Then
        n_MesIni = Val(IIf(s_Ano = s_AnoIni, cmbPeriodo(2).ListIndex, 1))
        n_MesFin = Val(IIf(s_Ano = s_AnoFin, cmbPeriodo(3).ListIndex, 12))
      End If
      ' Acumulación de saldos
      s_SaldoDeb = "": s_SaldoHab = ""
      For n_Index = n_MesIni To n_MesFin
        s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
        s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      Next n_Index
    
      ' Inserto los registros
      CadCrystal = "INSERT INTO " & ps_Prefijo & "trpRngEfiCco "
      If nFormato = "0" Then
        CadCrystal = CadCrystal & "SELECT a.codcta, MAX(" & Choose(gsIdioma, "b.detcta", "b.detctax") & ") AS detcta, "
      Else
        CadCrystal = CadCrystal & "SELECT a.codcco AS codcta, MAX(" & Choose(gsIdioma, "b.DetCCo", "b.DetCCox") & ") AS detcta, "
      End If
      ' primer paso: Acumulo las columnas por centro de costos o cuenta
      porsClone.MoveFirst
      Do While Not porsClone.EOF
        CadCrystal = CadCrystal & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
        CadCrystal = CadCrystal & "CASE a." & Choose(nFormato + 1, "codcco", "codcta") & " WHEN '" & porsClone!codcco & "' THEN "
        CadCrystal = CadCrystal & "(" & s_SaldoDeb & ")-(" & s_SaldoHab & ") "
        CadCrystal = CadCrystal & "ELSE 0 END, 0)), 2) AS " & Trim("x" & porsClone!numord) & ", "
        porsClone.MoveNext
      Loop
    
      ' segundo paso: Registro otros ultimo registro
      nRegistro = nRegistro - IIf(nRegistro > 1, 1, 0)
      For n_Index = nRegistro To 12
        CadCrystal = CadCrystal & "0 AS " & Trim("x" & Format(n_Index, "00")) & ", "
      Next n_Index
      nRegistro = nRegistro + IIf(nRegistro > 1, 1, 0)
      
      'Paso Adicional Agregado Septiembre del 2008
      CadCrystal = CadCrystal & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
      CadCrystal = CadCrystal & "IF(x.indpdocpr=1 AND LEFT(A.CODCCO,2)='" & Left(hoja, 2) & "', "
      CadCrystal = CadCrystal & "(" & s_SaldoDeb & ")-(" & s_SaldoHab & ") "
      CadCrystal = CadCrystal & ",0), 0)), 2) AS yYY , "
        
      ' tercer paso: genero le total general
      CadCrystal = CadCrystal & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
      CadCrystal = CadCrystal & "(" & s_SaldoDeb & ")-(" & s_SaldoHab & "), 0)), 2) AS xTotal "
      ' cuarto paso: seleccion de tablas y condicion
      CadCrystal = CadCrystal & "FROM (coccoacu a "
      
      If nFormato = "0" Then
        CadCrystal = CadCrystal & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.codcta=b.codcta) "
        CadCrystal = CadCrystal & "INNER JOIN cocco x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.codcco=x.codcco "
        CadCrystal = CadCrystal & "WHERE a.codemp='" & gsCodEmp & "' "
        CadCrystal = CadCrystal & "AND a.pdoano='" & s_Ano & "' "
        CadCrystal = CadCrystal & "AND (a.codcta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "') "
        CadCrystal = CadCrystal & "AND b.tpocta='" & TPOCTA_TRA & "' "
        CadCrystal = CadCrystal & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.codcco)=" & nNivel & " "
        CadCrystal = CadCrystal & "GROUP BY a.codcta "
        CadCrystal = CadCrystal & "ORDER BY a.codcta"
      Else
        CadCrystal = CadCrystal & "LEFT JOIN cocco b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCCo=b.CodCCo) "
        CadCrystal = CadCrystal & "INNER JOIN cocco x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.codcco=x.codcco "
        CadCrystal = CadCrystal & "WHERE a.codemp='" & gsCodEmp & "' "
        CadCrystal = CadCrystal & "AND a.pdoano='" & s_Ano & "' "
        CadCrystal = CadCrystal & "AND (a.codcco BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "') "
        CadCrystal = CadCrystal & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(b.CodCCo)=" & Right(gsNivCCo, 1) & " "
        CadCrystal = CadCrystal & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodCta)=" & nNivel & ""
        CadCrystal = CadCrystal & "GROUP BY a.codcco "
        CadCrystal = CadCrystal & "ORDER BY a.codcco"
      End If
      pocnnMain.Execute CadCrystal
    Next dnContador
  End If
  
     
  ' Obtengo los registros del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    CadCrystal = "SELECT codcta, detcta, "
    CadCrystal = CadCrystal & "ROUND(SUM(x00), 2) AS x00, "
    CadCrystal = CadCrystal & "ROUND(SUM(x01), 2) AS x01, "
    CadCrystal = CadCrystal & "ROUND(SUM(x02), 2) AS x02, "
    CadCrystal = CadCrystal & "ROUND(SUM(x03), 2) AS x03, "
    CadCrystal = CadCrystal & "ROUND(SUM(x04), 2) AS x04, "
    CadCrystal = CadCrystal & "ROUND(SUM(x05), 2) AS x05, "
    CadCrystal = CadCrystal & "ROUND(SUM(x06), 2) AS x06, "
    CadCrystal = CadCrystal & "ROUND(SUM(x07), 2) AS x07, "
    CadCrystal = CadCrystal & "ROUND(SUM(x08), 2) AS x08, "
    CadCrystal = CadCrystal & "ROUND(SUM(x09), 2) AS x09, "
    CadCrystal = CadCrystal & "ROUND(SUM(x10), 2) AS x10, "
    CadCrystal = CadCrystal & "ROUND(SUM(x11), 2) AS x11, "
    CadCrystal = CadCrystal & "ROUND(SUM(x12), 2) AS x12, "
    CadCrystal = CadCrystal & "ROUND(SUM(xXX), 2) AS xXX, "
    CadCrystal = CadCrystal & "ROUND(SUM(yYY), 2) AS yYY, "
    CadCrystal = CadCrystal & "ROUND(SUM(xTotal), 2) AS xTotal "
    CadCrystal = CadCrystal & "FROM " & ps_Prefijo & "trpRngEfiCco "
    CadCrystal = CadCrystal & "GROUP BY codcta, detcta "
    CadCrystal = CadCrystal & "ORDER BY codcta"
    .Source = CadCrystal
    .Open
  End With
  
'   Obtengo los registros del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    CadCrystal = "Insert Into " & ps_Prefijo & "trpRngEfiCcox "
    CadCrystal = CadCrystal & "SELECT codcta, detcta, "
    CadCrystal = CadCrystal & "ROUND(SUM(x00), 2) AS x00, "
    CadCrystal = CadCrystal & "ROUND(SUM(x01), 2) AS x01, "
    CadCrystal = CadCrystal & "ROUND(SUM(x02), 2) AS x02, "
    CadCrystal = CadCrystal & "ROUND(SUM(x03), 2) AS x03, "
    CadCrystal = CadCrystal & "ROUND(SUM(x04), 2) AS x04, "
    CadCrystal = CadCrystal & "ROUND(SUM(x05), 2) AS x05, "
    CadCrystal = CadCrystal & "ROUND(SUM(x06), 2) AS x06, "
    CadCrystal = CadCrystal & "ROUND(SUM(x07), 2) AS x07, "
    CadCrystal = CadCrystal & "ROUND(SUM(x08), 2) AS x08, "
    CadCrystal = CadCrystal & "ROUND(SUM(x09), 2) AS x09, "
    CadCrystal = CadCrystal & "ROUND(SUM(x10), 2) AS x10, "
    CadCrystal = CadCrystal & "ROUND(SUM(x11), 2) AS x11, "
    CadCrystal = CadCrystal & "ROUND(SUM(x12), 2) AS x12, "
    CadCrystal = CadCrystal & "ROUND(SUM(xXX), 2) AS xXX, "
    CadCrystal = CadCrystal & "ROUND(SUM(yYY), 2) AS yYY, "
    CadCrystal = CadCrystal & "ROUND(SUM(xTotal), 2) AS xTotal, "
    CadCrystal = CadCrystal & "ROUND(SUM(0), 2) AS aaa, "
    CadCrystal = CadCrystal & "'' as detagr, "
    CadCrystal = CadCrystal & "0 as bbb "
    CadCrystal = CadCrystal & "FROM " & ps_Prefijo & "trpRngEfiCco "
    CadCrystal = CadCrystal & "GROUP BY codcta, detcta "
    CadCrystal = CadCrystal & "ORDER BY codcta "
    .Source = CadCrystal
    .Open
  End With
  
  sql = " select codcta from " & ps_Prefijo & "trpRngEfiCcox"

  Rstfiltro.Open sql, pocnnMain, adOpenStatic, adLockOptimistic

  If Rstfiltro.RecordCount = 0 Then
     Exit Sub
  Else
     Rstfiltro.MoveFirst
      'recorre todo el recordset
     For contador1 = 0 To Rstfiltro.RecordCount - 1

        'sql = "SELECT sum(impcta_" & s_Moneda & ") FROM copdocprcta Where copdocprcta.codemp='" & gsCodEmp & "' and copdocprcta.pdoano='" & gsAnoAct & "' and left(copdocprcta.codcco,2)='" & hoja & "' and copdocprcta.codcta='" & Rstfiltro.Fields(0).Value & "' group by copdocprcta.codemp,copdocprcta.pdoano,copdocprcta.codcta,left(copdocprcta.codcco,2)"
        
        If chkRango.Value = vbChecked Then
            'sql = "SELECT sum(impcta_" & s_Moneda & ") FROM copdocprcta Where copdocprcta.codemp='" & gsCodEmp & "' and copdocprcta.pdoano between '" & s_AnoIni & "' and '" & s_AnoFin & "' and left(copdocprcta.codcco,2)='" & hoja & "' and copdocprcta.codcta='" & Rstfiltro.Fields(0).Value & "' and copdocprcta.mespvs between '" & s_Mes & "' and '" & s_Ano & "' group by copdocprcta.codemp,copdocprcta.pdoano,copdocprcta.codcta,left(copdocprcta.codcco,2)"
            sql = "SELECT sum(impcta_" & s_Moneda & ") FROM copdocprcta Where copdocprcta.codemp='" & gsCodEmp & "' and concat(copdocprcta.pdoano,copdocprcta.mespvs)>= '" & s_AnoIni & s_Mes & "' and concat(copdocprcta.pdoano,copdocprcta.mespvs) <= '" & s_AnoFin & Format(cmbPeriodo(3).ListIndex, "00") & "' and left(copdocprcta.codcco,2)='" & hoja & "' and copdocprcta.codcta='" & Rstfiltro.Fields(0).Value & "' group by copdocprcta.codemp,copdocprcta.codcta,left(copdocprcta.codcco,2)"
        Else
            sql = "SELECT sum(impcta_" & s_Moneda & ") FROM copdocprcta Where copdocprcta.codemp='" & gsCodEmp & "' and copdocprcta.pdoano='" & gsAnoAct & "' and left(copdocprcta.codcco,2)='" & hoja & "' and copdocprcta.codcta='" & Rstfiltro.Fields(0).Value & "' and copdocprcta.mespvs<='" & gsMesAct & " ' group by copdocprcta.codemp,copdocprcta.pdoano,copdocprcta.codcta,left(copdocprcta.codcco,2)"
        End If

        Rstdatos.Open sql, cnn, adOpenStatic, adLockOptimistic

        If Rstdatos.RecordCount = 0 Then
        Else
        Rstdatos.MoveFirst
        For contador2 = 0 To Rstdatos.RecordCount - 1

            With porstMRp
                If .State = adStateOpen Then .Close
                CadCrystal = "Update " & ps_Prefijo & "trpRngEfiCcox set " & ps_Prefijo & "trpRngEfiCcox.aaa = " & Rstdatos.Fields(0).Value & " where codcta='" & Rstfiltro.Fields(0).Value & "'"
                .Source = CadCrystal
                .Open
            End With

       Rstdatos.MoveNext
       Next
       End If
       Rstdatos.Close
    Rstfiltro.MoveNext
    Next
  End If
  Rstfiltro.Close
  
 Dim sentencia As String
 Dim sumatoria As String
 Dim desde As Integer
 
 'PRESUPUESTO
 
    sentencia = "sum("
    sumatoria = ""
    'Momentaneo
    If Right(cmbPeriodo(1).Text, 4) > Right(cmbPeriodo(0).Text, 4) Then
        For desde = 1 To 12
            sentencia = sentencia & "imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(12), "", "+")
            sumatoria = sumatoria & "p.imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(12), "", "+")
        Next
        sentencia = sentencia & ")"
    Else
        For desde = 1 To Int(gsMesAct)
            sentencia = sentencia & "imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(gsMesAct), "", "+")
            sumatoria = sumatoria & "p.imp" & s_Moneda & "_" & IIf(desde < 10, "0" & desde, desde) & IIf(desde = Int(gsMesAct), "", "+")
        Next
        sentencia = sentencia & ")"
    End If
    
 If gsMesAct <> "00" Then
    With porstMRp
       If .State = adStateOpen Then .Close
           If optAlcance(0).Value = True Then
               'CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.bbb=(select " & sentencia & " FROM copsp where codcta=x.codcta and left(codcco,2)='" & HOJA & "' and pdoano='" & s_Ano & "' and codemp='" & gsCodEmp & "' )  "
               CadCrystal = " UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.bbb=(select " & sentencia & " FROM copsp where codcta=x.codcta and left(codcco,2)='" & hoja & "' and pdoano>='" & Right(cmbPeriodo(0).Text, 4) & "' and pdoano<='" & Right(cmbPeriodo(1).Text, 4) & "' and codemp='" & gsCodEmp & "' )  "
            Else
               'CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.bbb=(select imp" & s_Moneda & "_" & gsMesAct & " FROM copsp where codcta=x.codcta and left(codcco,2)='" & HOJA & "' and pdoano='" & s_Ano & "' and codemp='" & gsCodEmp & "' )  "
               CadCrystal = " UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.bbb=(select imp" & s_Moneda & "_" & gsMesAct & " FROM copsp where codcta=x.codcta and left(codcco,2)='" & hoja & "' and pdoano>='" & Right(cmbPeriodo(0).Text, 4) & "' and pdoano<='" & Right(cmbPeriodo(1).Text, 4) & "' and codemp='" & gsCodEmp & "' )  "
            End If
       .Source = CadCrystal
       .Open
    End With
    With porstMRp
       If .State = adStateOpen Then .Close
           If optAlcance(0).Value = True Then
               'CadCrystal = "insert into trpRngEfiCcox select p.codcta,b.detcta,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,''," & sumatoria & " from copsp p inner join cocta b on p.codemp=b.codemp and p.pdoano=b.pdoano and p.codcta=b.codcta where p.codemp='" & gsCodEmp & "' and p.pdoano='" & s_Ano & "' and left(p.codcco,2)='" & hoja & "' and p.codcta not in (select x.codcta from trpRngEfiCcox x )"
               CadCrystal = " insert into trpRngEfiCcox select p.codcta,b.detcta,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,''," & sumatoria & " from copsp p inner join cocta b on p.codemp=b.codemp and p.pdoano=b.pdoano and p.codcta=b.codcta where p.codemp='" & gsCodEmp & "' and p.pdoano>='" & Right(cmbPeriodo(0).Text, 4) & "' and p.pdoano<='" & Right(cmbPeriodo(1).Text, 4) & "' and left(p.codcco,2)='" & hoja & "' and p.codcta not in (select x.codcta from trpRngEfiCcox x )"
            Else
               CadCrystal = " insert into trpRngEfiCcox select p.codcta,b.detcta,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'',imp" & s_Moneda & "_" & gsMesAct & " from copsp p inner join cocta b on p.codemp=b.codemp and p.pdoano=b.pdoano and p.codcta=b.codcta where p.codemp='" & gsCodEmp & "' and p.pdoano='" & s_Ano & "' and left(p.codcco,2)='" & hoja & "' and p.codcta not in (select x.codcta from trpRngEfiCcox x )"
           End If
       .Source = CadCrystal
       .Open
    End With
 End If
    
 
 With porstMRp
    If .State = adStateOpen Then .Close
    CadCrystal = "UPDATE " & ps_Prefijo & "trpRngEfiCcox x set x.detagr=(select detcta FROM cocta where codcta=left(x.codcta,2) and pdoano='" & s_Ano & "' and codemp='" & gsCodEmp & "' ) "
    .Source = CadCrystal
    .Open
 End With
      
  ' Obtengo los registros del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    CadCrystal = "SELECT codcta, detcta, "
    CadCrystal = CadCrystal & "if( bbb is null , 0 ,bbb ) as bbb, "
    CadCrystal = CadCrystal & "ROUND(SUM(x00), 2) AS x00, "
    CadCrystal = CadCrystal & "ROUND(SUM(x01), 2) AS x01, "
    CadCrystal = CadCrystal & "ROUND(SUM(x02), 2) AS x02, "
    CadCrystal = CadCrystal & "ROUND(SUM(x03), 2) AS x03, "
    CadCrystal = CadCrystal & "ROUND(SUM(x04), 2) AS x04, "
    CadCrystal = CadCrystal & "ROUND(SUM(x05), 2) AS x05, "
    CadCrystal = CadCrystal & "ROUND(SUM(x06), 2) AS x06, "
    CadCrystal = CadCrystal & "ROUND(SUM(x07), 2) AS x07, "
    CadCrystal = CadCrystal & "ROUND(SUM(x08), 2) AS x08, "
    CadCrystal = CadCrystal & "ROUND(SUM(x09), 2) AS x09, "
    CadCrystal = CadCrystal & "ROUND(SUM(x10), 2) AS x10, "
    CadCrystal = CadCrystal & "ROUND(SUM(x11), 2) AS x11, "
    CadCrystal = CadCrystal & "ROUND(SUM(x12), 2) AS x12, "
    CadCrystal = CadCrystal & "ROUND(SUM(0), 2) AS Suma, "
    CadCrystal = CadCrystal & "ROUND(SUM(aaa)-SUM(yyy), 2) AS SaldoComprometido "
    CadCrystal = CadCrystal & "FROM " & ps_Prefijo & "trpRngEfiCcox "
    CadCrystal = CadCrystal & "GROUP BY codcta, detcta "
    If Valx = False Then
        CadCrystal = CadCrystal & "HAVING (x00+x01+x02+x03+x04+x05+x06+x07+x08+x09+x10+bbb) <> 0.00 OR ROUND(SUM(aaa), 2) <> 0 "
    Else
        CadCrystal = CadCrystal & "HAVING (x00+x01+x02+x03+x04+x05+x06+x07+x08+x09+x10+bbb) >= 0.00 OR ROUND(SUM(aaa), 2) >= 0 "
    End If
    CadCrystal = CadCrystal & ""
    .Source = CadCrystal
    .Open
  End With
  
  'CadCrystal = CadCrystal & " Union all select left(codcta,2),detagr,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0 from trpRngEfiCcox "
  'CadCrystal = CadCrystal & " GROUP BY left(codcta,2), detagr "
  'CadCrystal = CadCrystal & " ORDER BY codcta"
  
    CadCrystal = CadCrystal & " Union all "
    CadCrystal = CadCrystal & "SELECT left(codcta,2), detagr, "
    CadCrystal = CadCrystal & "if( bbb is null , 0 ,bbb ) as bbb, "
    CadCrystal = CadCrystal & "ROUND(SUM(x00), 2) AS x00, "
    CadCrystal = CadCrystal & "ROUND(SUM(x01), 2) AS x01, "
    CadCrystal = CadCrystal & "ROUND(SUM(x02), 2) AS x02, "
    CadCrystal = CadCrystal & "ROUND(SUM(x03), 2) AS x03, "
    CadCrystal = CadCrystal & "ROUND(SUM(x04), 2) AS x04, "
    CadCrystal = CadCrystal & "ROUND(SUM(x05), 2) AS x05, "
    CadCrystal = CadCrystal & "ROUND(SUM(x06), 2) AS x06, "
    CadCrystal = CadCrystal & "ROUND(SUM(x07), 2) AS x07, "
    CadCrystal = CadCrystal & "ROUND(SUM(x08), 2) AS x08, "
    CadCrystal = CadCrystal & "ROUND(SUM(x09), 2) AS x09, "
    CadCrystal = CadCrystal & "ROUND(SUM(x10), 2) AS x10, "
    CadCrystal = CadCrystal & "ROUND(SUM(x11), 2) AS x11, "
    CadCrystal = CadCrystal & "ROUND(SUM(x12), 2) AS x12, "
    CadCrystal = CadCrystal & "ROUND(SUM(0), 2) AS Suma, "
    CadCrystal = CadCrystal & "ROUND(SUM(aaa)-SUM(yyy), 2) AS SaldoComprometido "
    CadCrystal = CadCrystal & "FROM " & ps_Prefijo & "trpRngEfiCcox "
    CadCrystal = CadCrystal & "GROUP BY left(codcta,2), detagr "
    If Valx = False Then
        CadCrystal = CadCrystal & "HAVING (x00+x01+x02+x03+x04+x05+x06+x07+x08+x09+x10+bbb) <> 0.00 OR ROUND(SUM(aaa), 2) <> 0 "
    Else
        CadCrystal = CadCrystal & "HAVING (x00+x01+x02+x03+x04+x05+x06+x07+x08+x09+x10+bbb) >= 0.00 OR ROUND(SUM(aaa), 2) >= 0 "
    End If
    CadCrystal = CadCrystal & " ORDER BY codcta"
  
  strsql = CadCrystal
  
  cols = 18
  Dim rsexportar As New Recordset
  rsexportar.Open strsql, cnn, adOpenStatic, adLockOptimistic
  On Error GoTo error
  rsexportar.MoveFirst
  For i = 1 To rsexportar.RecordCount
    If i = 1 Then
        For j = 1 To cols
            ApExcel.Cells(i + 3, j).formula = rsexportar.Fields(j - 1).Name
        Next
    End If
    For j = 1 To cols
    ApExcel.Cells(i + 4, j).formula = rsexportar(j - 1)
    Next
  rsexportar.MoveNext
  Next
  
  porsClone.Close
  Set porsClone = Nothing
  ' Elimino y genero el archivo del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCco", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCco') DROP TABLE #trpRngEfiCco")
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngEfiCcox", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngEfiCcox') DROP TABLE #trpRngEfiCcox")
  ppHabilitacion True
  
error:
End Sub
Sub darformato(hoja As String)
Dim rscostos As New ADODB.Recordset
Dim i As Integer
Dim k As Integer
Dim columna As Integer
Dim ultimacolumna As Integer
Dim sql As String
Dim contar As Integer
Dim ultimov As Integer

contar = 0

ReDim ACostos(20)
ReDim NCostos(20)

i = 1
sql = "select codcco,detcco from cocco where pdoano='" & Right(cmbPeriodo(1), 4) & "' and codemp='" & gsCodEmp & "' and length(codcco)=5 and left(codcco,2)='" & hoja & "'"

On Error GoTo error
rscostos.Open sql, cnn, adOpenDynamic, adLockPessimistic
Do Until rscostos.EOF
    ACostos(i - 1) = rscostos(0).Value
    NCostos(i - 1) = rscostos(1).Value
    i = i + 1
rscostos.MoveNext
Loop
rscostos.Close

ApExcel.Cells(4, 1).Select
ApExcel.ActiveCell = "Cuenta"
ApExcel.Cells(4, 2).Select
ApExcel.ActiveCell = "Detalle"
ApExcel.Cells(4, 3).Select
ApExcel.ActiveCell = "Presupuesto"

columna = 4

For k = LBound(ACostos) To UBound(ACostos)
    If ACostos(k) <> "" Then
    
    ApExcel.Cells(3, columna).Select
    ApExcel.ActiveCell = ACostos(k)
    
    ApExcel.Cells(4, columna).Select
    ApExcel.ActiveCell = NCostos(k)
            
    columna = columna + 1
    
    End If
Next

ultimacolumna = columna

Select Case columna
Case 4
    ApExcel.Columns("H:P").Select
    ApExcel.Selection.Delete
    Exit Sub
Case 11
    ApExcel.Columns("K:P").Select
    ApExcel.Selection.Delete
Case 12
    ApExcel.Columns("L:P").Select
    ApExcel.Selection.Delete
End Select

Dim fila As Integer
fila = 5
Do While True
    If IsEmpty(ApExcel.Cells(fila, 1)) Then Exit Do
    
    mensaje.Caption = "Procesando Cuenta Contable " & ApExcel.Cells(fila, 1).Value & " del Formato " & hoja
    mensaje.Refresh
    
    ApExcel.Cells(fila, columna).Select
    ApExcel.ActiveCell.FormulaR1C1 = "=SUM(RC[-" & (columna - 4) & "]:RC[-1])"
    ApExcel.Cells(fila, columna + 3).Select
    ApExcel.ActiveCell.FormulaR1C1 = "=SUM(RC[-" & (3) & "]:RC[-1])"
    ApExcel.Cells(fila, columna + 5).Select
    ApExcel.ActiveCell.FormulaR1C1 = "=SUM(RC[-" & (2) & "]:RC[-1])"
    fila = fila + 1
Loop

ApExcel.Cells(4, columna + 2).Select
ApExcel.ActiveCell = "Pendiente"
ApExcel.Cells(4, columna + 3).Select
ApExcel.ActiveCell = "Total Peru"
ApExcel.Cells(4, columna + 4).Select
ApExcel.ActiveCell = "España"
ApExcel.Cells(4, columna + 5).Select
ApExcel.ActiveCell = "Total Proyecto"
'************************************************************************************
Dim valoresA() As Integer
Dim valoresB() As Integer

Dim b As Boolean
ReDim valoresA(200)
ReDim valoresB(200)
Dim ca As Integer
Dim cb As Integer
Dim divi As Integer

i = 0
ca = 0
cb = 0

b = False

For i = 6 To fila
        
        If Len(ApExcel.Cells(i, 1).Value) = 8 And b = False Then
            valoresA(ca) = i
            b = True
            ca = ca + 1
        End If
     
        If Len(ApExcel.Cells(i, 1).Value) <> 8 And b = True Then
            valoresB(cb) = i - 1
            b = False
            cb = cb + 1
        End If
        
        If Left(ApExcel.Cells(i, 1).Value, 1) >= 0 And Left(ApExcel.Cells(i, 1).Value, 1) <= 7 Then
           divi = i
        End If
        
Next

For i = LBound(valoresA) To UBound(valoresA)
    If valoresA(i) <> 0 Then
        
        If valoresB(i) <> 0 Then
        
         ApExcel.Range("C" & valoresA(i) - 1 & "").Select
         'ApExcel.ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[" & (valoresB(i) - valoresA(i)) + 1 & "]C)"
         ApExcel.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[1]C:R[" & (valoresB(i) - valoresA(i)) + 1 & "]C)"
    
         ApExcel.Range("C" & valoresA(i) - 1 & "").Select
         ApExcel.Selection.AutoFill Destination:=ApExcel.Range(ApExcel.Cells(valoresA(i) - 1, 3), ApExcel.Cells(valoresA(i) - 1, ultimacolumna + 5))
         ApExcel.Range(ApExcel.Cells(valoresA(i) - 1, 3), ApExcel.Cells(valoresA(i) - 1, ultimacolumna + 5)).Select
         
        Else
        End If
       
    End If
Next


ApExcel.Range(ApExcel.Cells(1, 4), ApExcel.Cells(1, ultimacolumna - 1)).Select
ApExcel.Selection.Columns.Group


For i = LBound(valoresA) To UBound(valoresA)
    If valoresA(i) <> 0 Then
        If valoresB(i) <> 0 Then
        ApExcel.Range(ApExcel.Cells(valoresA(i), 1), ApExcel.Cells(valoresB(i), 1)).Select
        Else
        'ApExcel.Range(ApExcel.Cells(valoresA(i), 1), ApExcel.Cells(Contador, 1)).Select
        End If
        ApExcel.Selection.Rows.Group
    End If
Next

'*******************************************************************************************

Dim formula As String

fila = 6
Do While True
    If IsEmpty(ApExcel.Cells(fila, 1)) Then Exit Do
        fila = fila + 1
Loop

If divi > 0 Then

ApExcel.Rows("" & divi + 1 & ":" & divi + 1 & "").Select
ApExcel.Selection.Insert 'Shift:=ApExcel.xlDown

For i = LBound(valoresA) To UBound(valoresA)
    If valoresA(i) <> 0 Then
        If valoresB(i) <> 0 Then
        
        If valoresA(i) <= divi Then formula = formula & "R[-" & divi - valoresA(i) + 2 & "]C+"
         
        Else
        End If
       
    End If
Next
 
formula = "=" & Left(formula, Len(formula) - 1)

ApExcel.Range("C" & divi + 1).Select
ApExcel.ActiveCell.FormulaR1C1 = formula

ApExcel.Range("C" & divi + 1 & "").Select
ApExcel.Selection.AutoFill Destination:=ApExcel.Range(ApExcel.Cells(divi + 1, 3), ApExcel.Cells(divi + 1, ultimacolumna + 5))
ApExcel.Range(ApExcel.Cells(divi + 1, 3), ApExcel.Cells(divi + 1, ultimacolumna + 6)).Select

Dim xvalor As Integer

formula = ""

For i = LBound(valoresA) To UBound(valoresA)
    If valoresA(i) <> 0 Then
        If valoresB(i) <> 0 Then
        
        'If valoresA(i) >= divi Then formula = formula & "R[-" & fila - valoresA(i) + 2 & "]C+"
        'If valoresA(i) >= divi Then formula = formula & "R[-" & fila - valoresA(i) + 2 & "]C,"
        
        If valoresA(i) < divi Then xvalor = valoresA(i)
        
        If valoresA(i) >= divi Then
          
        If divi = xvalor Then
            If contar > 1 Then formula = formula & "R[-" & fila - valoresA(i - 1) + 1 & "]C:" & "R[-" & fila - valoresA(i) + 3 & "]C,"
        Else
            If contar > 0 Then formula = formula & "R[-" & fila - valoresA(i - 1) + 1 & "]C:" & "R[-" & fila - valoresA(i) + 3 & "]C,"
        End If
        
        
        contar = contar + 1
        ultimov = valoresA(i)
        End If
        
        End If
    
    End If
Next


  
'formula = "=" & Left(formula, Len(formula) - 1)
formula = formula & "R[-" & fila - ultimov + 1 & "]C:" & "R[-" & 2 & "]C,"

formula = "=SUBTOTAL(9," & Left(formula, Len(formula) - 1) & ")"

ApExcel.Range("C" & fila + 2).Select
ApExcel.ActiveCell.FormulaR1C1 = formula

ApExcel.Range("C" & fila + 2).Select
ApExcel.Selection.AutoFill Destination:=ApExcel.Range(ApExcel.Cells(fila + 2, 3), ApExcel.Cells(fila + 2, ultimacolumna + 5))
ApExcel.Range(ApExcel.Cells(fila + 2, 3), ApExcel.Cells(fila + 2, ultimacolumna + 6)).Select

ApExcel.Range("C" & fila + 3).Select
ApExcel.ActiveCell.FormulaR1C1 = "=R[-" & (fila + 2) - divi & "]C+R[-1]C"
'ApExcel.ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[-" & (fila + 2) - divi & "]C,R[-1]C)"

ApExcel.Range("C" & fila + 3).Select
ApExcel.Selection.AutoFill Destination:=ApExcel.Range(ApExcel.Cells(fila + 3, 3), ApExcel.Cells(fila + 3, ultimacolumna + 5))
ApExcel.Range(ApExcel.Cells(fila + 3, 3), ApExcel.Cells(fila + 3, ultimacolumna + 6)).Select

Else

For i = LBound(valoresA) To UBound(valoresA)
    If valoresA(i) <> 0 Then
        If valoresB(i) <> 0 Then
        
        formula = formula & "R[-" & fila - valoresA(i) + 2 & "]C+"
         
        Else
        End If
       
    End If
Next
 
formula = "=" & Left(formula, Len(formula) - 1)

ApExcel.Range("C" & fila + 1).Select
ApExcel.ActiveCell.FormulaR1C1 = formula

ApExcel.Range("C" & fila + 1).Select
ApExcel.Selection.AutoFill Destination:=ApExcel.Range(ApExcel.Cells(fila + 1, 3), ApExcel.Cells(fila + 1, ultimacolumna + 5))
ApExcel.Range(ApExcel.Cells(fila + 1, 3), ApExcel.Cells(fila + 1, ultimacolumna + 6)).Select

End If


error:
End Sub
