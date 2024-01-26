VERSION 5.00
Begin VB.Form frmRBceCpbCCo 
   Caption         =   "[título]"
   ClientHeight    =   6690
   ClientLeft      =   1620
   ClientTop       =   1395
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6990
   Begin VB.CheckBox ChkImpCtas 
      Caption         =   "Imp.Tdas.Ctas"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5625
      TabIndex        =   52
      Top             =   3075
      Width           =   1395
   End
   Begin VB.CheckBox chkRango 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1410
      TabIndex        =   44
      Top             =   4920
      Width           =   180
   End
   Begin VB.Frame fraRngPeriodo 
      Caption         =   " Rango Periodos "
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   20
      TabIndex        =   45
      Top             =   4920
      Width           =   4215
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   51
         Text            =   "Mes Final"
         Top             =   645
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   50
         Text            =   "Mes Inicio"
         Top             =   300
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   855
         TabIndex        =   49
         Text            =   "Año Final"
         Top             =   645
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   47
         Text            =   "Año Inicio"
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   46
         Top             =   345
         Width           =   750
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   48
         Top             =   690
         Width           =   750
      End
   End
   Begin VB.Frame fraNivelCenCos 
      Caption         =   " Nivel de Centro de Costo "
      ForeColor       =   &H80000002&
      Height          =   840
      Left            =   -30
      TabIndex        =   38
      Top             =   4035
      Width           =   4335
      Begin VB.OptionButton optNivCCo 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   3
         Left            =   3000
         TabIndex        =   43
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   41
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   2
         Left            =   2040
         TabIndex        =   42
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5625
      TabIndex        =   11
      Top             =   2790
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   35
      Top             =   4110
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   37
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   36
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraNivelCuenta 
      Caption         =   "Nivel de Cuentas"
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   3180
      Width           =   6975
      Begin VB.OptionButton optNivCta 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   7
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "8 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   6
         Left            =   5880
         TabIndex        =   19
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "7 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   5
         Left            =   4920
         TabIndex        =   18
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "6 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   4
         Left            =   3960
         TabIndex        =   17
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   3
         Left            =   3000
         TabIndex        =   16
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   2
         Left            =   2040
         TabIndex        =   15
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCta 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Frame fraAlcance 
      Caption         =   "Alcance"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   20
      TabIndex        =   28
      Top             =   2400
      Width           =   2460
      Begin VB.OptionButton optAlcance 
         Caption         =   "del mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   255
         Width           =   1080
      End
      Begin VB.OptionButton optAlcance 
         Caption         =   "al mes"
         ForeColor       =   &H80000001&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   915
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2430
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rangos"
      ForeColor       =   &H00800000&
      Height          =   2265
      Left            =   -15
      TabIndex        =   21
      Top             =   60
      Width           =   6990
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRBceCpbCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1860
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRBceCpbCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1500
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
         TabIndex        =   7
         Top             =   1845
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
         TabIndex        =   6
         Top             =   1485
         Width           =   945
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   4500
         Picture         =   "frmRBceCpbCCo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   630
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
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   4500
         Picture         =   "frmRBceCpbCCo.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   1065
         TabIndex        =   31
         Top             =   1845
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
         TabIndex        =   30
         Top             =   1500
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   1260
         Width           =   585
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centros de Costo"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   1215
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
         Index           =   3
         Left            =   780
         TabIndex        =   25
         Top             =   855
         Width           =   3720
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
         Left            =   780
         TabIndex        =   24
         Top             =   495
         Width           =   3720
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
      ScaleWidth      =   6990
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6150
      Width           =   6990
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
         Picture         =   "frmRBceCpbCCo.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   1245
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
         Picture         =   "frmRBceCpbCCo.frx":0BDA
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Height          =   495
         Index           =   1
         Left            =   1245
         Picture         =   "frmRBceCpbCCo.frx":0D24
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   2
      Left            =   4965
      TabIndex        =   9
      Top             =   2490
      Width           =   690
   End
End
Attribute VB_Name = "frmRBceCpbCCo"
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
Private porstCOCta As ADODB.Recordset
Private porstCoCCo As ADODB.Recordset
Private pnNivCta As Byte
']
Private Sub chkRango_Click()
  fraRngPeriodo.Enabled = (chkRango.Value = vbChecked)
End Sub
Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstCoCCo = New ADODB.Recordset
   
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
   With porstCOCta
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
      .Source = .Source & "FROM CoCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstCoCCo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS DetCCo "
      .Source = .Source & "FROM CoCCo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      For dnContador = 2 To 3
         .Item(dnContador).DataField = "CodCCo"
         .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Centro de Costo :", "Cuentas :", "Moneda :", "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Cost Center :", "Accounts :", "Currency :", "Beginning :", "End :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraAlcance.Caption = Choose(gsIdioma, "Alcance", "Scope")
  optAlcance(0).Caption = Choose(gsIdioma, "al mes", "to month")
  optAlcance(1).Caption = Choose(gsIdioma, "del mes", "from month")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraNivelCuenta.Caption = Choose(gsIdioma, "Nivel de Cuentas", "Account Level")
  optNivCta(7).Caption = Choose(gsIdioma, "Detalle", "Detail")
  optNivCta(0).Caption = Choose(gsIdioma, "2 dígitos", "2 digits")
  optNivCta(1).Caption = Choose(gsIdioma, "3 dígitos", "3 digits")
  optNivCta(2).Caption = Choose(gsIdioma, "4 dígitos", "4 digits")
  optNivCta(3).Caption = Choose(gsIdioma, "5 dígitos", "5 digits")
  optNivCta(4).Caption = Choose(gsIdioma, "6 dígitos", "6 digits")
  optNivCta(5).Caption = Choose(gsIdioma, "7 dígitos", "7 digits")
  optNivCta(6).Caption = Choose(gsIdioma, "8 dígitos", "8 digits")
  fraNivelCenCos.Caption = Choose(gsIdioma, "Nivel Centro Costos", "Cost Center Level")
  optNivCCo(4).Caption = Choose(gsIdioma, "Detalle", "Detail")
  optNivCCo(0).Caption = Choose(gsIdioma, "2 dígitos", "2 digits")
  optNivCCo(1).Caption = Choose(gsIdioma, "3 dígitos", "3 digits")
  optNivCCo(2).Caption = Choose(gsIdioma, "4 dígitos", "4 digits")
  optNivCCo(3).Caption = Choose(gsIdioma, "5 dígitos", "5 digits")
  fraRngPeriodo.Caption = Choose(gsIdioma, "Rango Periodos", "Range of Periods")
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
    
    optAlcance(0).Value = True
    For dnContador = 1 To Len(gsNivCta)
        optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Visible = True
        Select Case dnContador
         Case Is = 1
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 120
         Case Is = 2
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 1080
         Case Is = 3
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 2040
         Case Is = 4
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 3000
         Case Is = 5
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 3960
         Case Is = 6
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 4920
         Case Is = 7
            optNivCta(Val(Mid(gsNivCta, dnContador, 1)) - 2).Left = 5880
        End Select
    Next
'    optNivCta(Val(Mid(gsNivCta, dncontador - 1, 1)) - 2).Value = True
    optNivCta(7).Value = True
    pnNivCta = 9
    fraNivelCuenta.Width = optNivCta(Val(Mid(gsNivCta, dnContador - 1, 1)) - 2).Left + 1035

    For dnContador = 1 To Len(gsNivCCo)
        optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Visible = True
        Select Case dnContador
         Case Is = 1
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 120
         Case Is = 2
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 1080
         Case Is = 3
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 2040
         Case Is = 4
            optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 3000
        End Select
    Next
    optNivCCo(4).Value = True
    fraNivelCenCos.Width = optNivCCo(Val(Mid(gsNivCCo, dnContador - 1, 1)) - 2).Left + 1300

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !codcta
      .MoveFirst
      txtDato(0).Text = !codcta
   End With
   With porstCoCCo
      .MoveLast
      txtDato(3).Text = !codcco
      .MoveFirst
      txtDato(2).Text = !codcco
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
   If txtDato(3).Text <> "" Then ppAyuDet 3
   
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
   porstCoCCo.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstCoCCo = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2, 3
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    
  Dim dnContador As Integer, n_Index As Integer
  Dim s_Sentencia As String, s_Sql As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_Moneda As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_NivCoCCo As Integer, n_MesIni As Integer, n_MesFin As Integer
  Dim l_CreateTB As Boolean
  
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
  
  ppHabilitacion False
    
  n_NivCoCCo = IIf(optNivCCo(0).Value, 2, IIf(optNivCCo(1).Value, 3, 5))
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
    
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCco", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCco') DROP TABLE #trpRngBceCco")
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    n_MesIni = Val(IIf(optAlcance(0).Value, 0, gsMesAct))
    n_MesFin = Val(gsMesAct)
    If chkRango.Value = vbChecked Then
      n_MesIni = Val(IIf(s_Ano = s_AnoIni, cmbPeriodo(2).ListIndex, 1))
      n_MesFin = Val(IIf(s_Ano = s_AnoFin, cmbPeriodo(3).ListIndex, 13))
    End If
    ' Acumulación de saldos
    s_SaldoDeb = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
    s_SaldoHab = "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "("
    For n_Index = n_MesIni To n_MesFin
      s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
    Next n_Index
    s_SaldoDeb = s_SaldoDeb & ", 0), 2) AS cSumaD, "
    s_SaldoHab = s_SaldoHab & ", 0), 2) AS cSumaH "
    
    ' Registros iniciales de saldos
    s_Sentencia = "SELECT a.CodCCo AS CodCCo, " & Choose(gsIdioma, "b.DetCCo", "b.DetCCox") & " AS DetCCo, a.CodCta, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, c.TpoSdo, c.TpoCta, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & s_SaldoHab
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCco ", "")
    s_Sentencia = s_Sentencia & "FROM ((COCCoAcu a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCCo=b.CodCCo) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND a.CodCCo BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    If pnNivCta = 2 Then
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    Else
      If pnNivCta = 9 Then
        s_Sentencia = s_Sentencia & "AND c.TpoCta='" & TPOCTA_TRA & "' "
      Else
        s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
        s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<" & pnNivCta & " AND c.TpoCta= " & TPOCTA_TRA & ")) "
      End If
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCCo))=" & n_NivCoCCo & " "
    s_Sentencia = s_Sentencia & "ORDER BY a.CodCCo, a.CodCta"
    ' Executo la sentencia
    If Not l_CreateTB Then
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trpRngBceCco ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trpRngBceCco "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
   
  With porstMRp
    If .State = adStateOpen Then .Close
    s_Sentencia = "SELECT CodCCo, DetCCo, CodCta, DetCta, TpoSdo, TpoCta, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaH), 2) AS cSumaH "
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCco "
    s_Sentencia = s_Sentencia & "GROUP BY CodCCo, CodCta, DetCCo, DetCta, TpoSdo, TpoCta "
    If ChkImpCtas = 0 Then
           s_Sentencia = s_Sentencia & "HAVING (" & IIf(ps_Plataforma = pSrvMySql, "cSumaD+cSumaH", "ROUND(SUM(cSumaD), 2)+ROUND(SUM(cSumaH), 2)") & "<> 0.00) "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY CodCCo, CodCta"
    .Source = s_Sentencia
    .Open
  End With
     
  s_Sentencia = IIf(chkRango.Value = vbChecked, cmbPeriodo(2).Text & " - " & cmbPeriodo(0).Text, "")
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptRBceCpbCCo.rpt"
      'Fórmular propias.
      .Formulas(5) = "mPeriodo='" & s_Sentencia & " " & IIf(optAlcance(0).Value, Choose(gsIdioma, "Acumulado - ", "Accrued - "), "") & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
      '.Formulas(6) = "fSaldoInicial=" & gsIniMesCnt
      ']
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
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRBceCpbCCo.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = IIf(optAlcance(0).Value, "A ", "") & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
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
  ' elimino el archivo temporal
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCco", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCco') DROP TABLE #trpRngBceCco")
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

Private Sub optNivCta_Click(Index As Integer)

    pnNivCta = Index + 2
    
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
   Case 0, 1, 2, 3                   'Cambiar (añadir índices).
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
   Case 2, 3                           'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
    Case 2, 3
       If txtDato(tnIndex).Text = "" Then
          lblDatoDeta(tnIndex).Caption = ""
          Exit Function
       End If
       With porstCoCCo
          .MoveFirst
          .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
             MsgBox TEXT_8006, vbExclamation
             ppAyuDet = True
          Else
             lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetCCo), "", !DetCCo)
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
