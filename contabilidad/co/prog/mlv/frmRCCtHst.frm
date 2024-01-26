VERSION 5.00
Begin VB.Form frmRCCtHst 
   Caption         =   "[título]"
   ClientHeight    =   5385
   ClientLeft      =   2460
   ClientTop       =   2115
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7290
   Begin VB.ComboBox cboInformacion 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3345
      Width           =   1830
   End
   Begin VB.CheckBox chkRango 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1470
      TabIndex        =   22
      Top             =   3720
      Width           =   180
   End
   Begin VB.Frame fraRngPeriodo 
      Caption         =   " Rango Periodos "
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   90
      TabIndex        =   21
      Top             =   3705
      Width           =   4215
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   28
         Text            =   "Mes Final"
         Top             =   645
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   25
         Text            =   "Mes Inicio"
         Top             =   300
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   855
         TabIndex        =   27
         Text            =   "Año Final"
         Top             =   645
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   24
         Text            =   "Año Inicio"
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   345
         Width           =   765
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   26
         Top             =   690
         Width           =   765
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   33
      Top             =   3750
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   35
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   34
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Auxiliar"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   15
      Top             =   2490
      Width           =   7290
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
         TabIndex        =   16
         Top             =   315
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   6885
         Picture         =   "frmRCCtHst.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   325
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
         Left            =   1365
         TabIndex        =   17
         Top             =   315
         Width           =   5520
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   90
      Width           =   7290
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
         Left            =   150
         TabIndex        =   11
         Top             =   1515
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
         Index           =   4
         Left            =   150
         TabIndex        =   13
         Top             =   1875
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   6600
         Picture         =   "frmRCCtHst.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1515
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   4
         Left            =   6600
         Picture         =   "frmRCCtHst.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1875
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6600
         Picture         =   "frmRCCtHst.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   855
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6600
         Picture         =   "frmRCCtHst.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   495
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
         Left            =   150
         TabIndex        =   8
         Top             =   840
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
         Left            =   150
         TabIndex        =   6
         Top             =   480
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
         Index           =   3
         Left            =   780
         TabIndex        =   12
         Top             =   1515
         Width           =   5850
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
         Index           =   4
         Left            =   780
         TabIndex        =   14
         Top             =   1875
         Width           =   5850
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costos"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Top             =   1245
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
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   840
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
         Left            =   1080
         TabIndex        =   7
         Top             =   495
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   585
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
      ScaleWidth      =   7290
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4845
      Width           =   7290
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
         Picture         =   "frmRCCtHst.frx":0852
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
         Picture         =   "frmRCCtHst.frx":099C
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
         Picture         =   "frmRCCtHst.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Informacion : "
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3405
      Width           =   960
   End
End
Attribute VB_Name = "frmRCCtHst"
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
Private porstTGAux As ADODB.Recordset
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
   Set porstTGAux = New ADODB.Recordset
   
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
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
  With porstCoCCo
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS detcco "
    .Source = .Source & "FROM CoCCo "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodCCo"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TGAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With

 ']

 '[Parámetros.                         'Cambiar.
  With cboInformacion
    .AddItem "General", 0
    .AddItem "Pendiente", 1
    .AddItem "Cancelados", 2
  End With
   
  With txtDato
    For dnContador = 0 To 1
      .Item(dnContador).DataField = "CodCta"
      .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
    Next
    For dnContador = 3 To 4
      .Item(dnContador).DataField = "codcco"
      .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
    Next
    .Item(2).DataField = "codaux"
    .Item(2).MaxLength = porstTGAux.Fields(.Item(2).DataField).DefinedSize
  End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Información :", "Inicio :", "Fin :", "Centro de Costos :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Information :", "Beginning :", "End :", "Cost Center :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraAuxiliar.Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraRngPeriodo.Caption = Choose(gsIdioma, "Rango Periodos", "Range of Periods")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
    'With cboTpoMon
    '    .AddItem TPOMON_NAC_TXT_1, 0
    '    .AddItem TPOMON_EXT_TXT_1, 1
    'End With
    'cboTpoMon.ListIndex = TPOMON_NAC_IND

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
    .MoveLast
    txtDato(1).Text = !codcta
    .MoveFirst
    txtDato(0).Text = !codcta
   End With
   
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
   If txtDato(3).Text <> "" Then ppAyuDet 3
   If txtDato(4).Text <> "" Then ppAyuDet 4
  
  'Otros.
   
  'Características de impresión.
   cboInformacion.ListIndex = 0
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
   porstTGAux.Close
   porstCOCta.Close
   porstCoCCo.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstCoCCo = Nothing
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2, 3, 4
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim dnContador As Integer, s_Sql As String
  Dim s_Sentencia As String, s_Catalogo As String, s_Comparar As String
  Dim s_AnoIni As String, s_AnoFin As String, sQuiebre As String
  Dim s_Ano As String, s_Mes As String
  Dim l_CreateTB As Boolean
       
  sQuiebre = "N"
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial", "End month must be equal or more than opening"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
   
  ppHabilitacion False
    
  ' Elimino y genero el archivo temporal del reporte
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#trptcctctahs_') DROP TABLE #trptcctctahs"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptcctctahs", s_Sentencia)
  ' Inserto los registros
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    s_Sentencia = "SELECT a.pdoano AS cAno, a.MesPvs, a.CodCta, a.CodAux, a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, "
    If (Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> "") Then
      s_Sentencia = s_Sentencia & "a.codcco, " & Choose(gsIdioma, "e.detcco", "e.detccox") & " AS detcco, "
    Else
      s_Sentencia = s_Sentencia & "Null AS codcco, Null AS detcco, "
    End If
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocum, "
    s_Sentencia = s_Sentencia & "a.FehOpe, a.FeEDoc, a.FeVDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, b.RazAux, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) AS cDebeMN, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) AS cHaberMN, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END) AS cDebeME, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) AS cHaberME "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trptcctctahs ", "")
    s_Sentencia = s_Sentencia & "FROM ((((COCpbDet a "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN Cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.Codcta=d.Codcta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.codcco=e.codcco) "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    s_Sentencia = s_Sentencia & "AND LEFT(a.codcta, " & Len(Trim(txtDato(0).Text)) & ")>='" & txtDato(0).Text & "' "
    s_Sentencia = s_Sentencia & "AND LEFT(a.codcta, " & Len(Trim(txtDato(1).Text)) & ")<='" & txtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND (a.ImpMN<> 0.00 OR a.ImpME<> 0.00) "
    ' Si activo el rango de periodos
    If chkRango.Value = vbChecked Then
      s_Mes = Format(IIf(s_Ano = s_AnoIni, cmbPeriodo(2).ListIndex, "1"), "00")
      s_Sentencia = s_Sentencia & "AND a.Mespvs >='" & s_Mes & "' "
      If (s_Ano = s_AnoFin) Then
        s_Mes = Format(cmbPeriodo(3).ListIndex, "00")
        s_Sentencia = s_Sentencia & "AND a.Mespvs <='" & s_Mes & "' "
      End If
    Else
      s_Sentencia = s_Sentencia & "AND a.Mespvs <='" & gsMesAct & "' "
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '') <>'' AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTDc, '') <>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.SerDoc, '') <>'' AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.NroDoc, '') <>'' AND d.inddoc='1' "
    If Trim(txtDato(2).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.CodAux='" & txtDato(2).Text & "' "
    End If
    If (Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> "") Then
      s_Sentencia = s_Sentencia & "AND LEFT(a.codcco, " & Len(Trim(txtDato(3).Text)) & ")>='" & txtDato(3).Text & "' "
      s_Sentencia = s_Sentencia & "AND LEFT(a.codcco, " & Len(Trim(txtDato(4).Text)) & ")<='" & txtDato(4).Text & "' "
      sQuiebre = "S"
    End If
    s_Sentencia = s_Sentencia & "ORDER BY " & IIf((Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> ""), " a.codcco, ", "") & "a.codcta, a.codaux, a.codtdc, a.serdoc, a.NroDoc, a.TpoPvs, a.MesPvs, a.FehOpe"
    
    ' Executo la sentencia
    If Not l_CreateTB Then
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptcctctahs ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trptcctctahs "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
   
  ' Obtengo la informacion pendiente o cancelada
  If cboInformacion.ListIndex <> 0 Then
    s_Comparar = Choose(cboInformacion.ListIndex, "<>", "=")
    s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpfiltros_') DROP TABLE #tmpfiltros"
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpfiltros", s_Sentencia)
  
    s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpfiltros ", "")
    s_Sentencia = s_Sentencia & "SELECT " & IIf((Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> ""), "a.codcco, ", "") & "a.codcta, a.codaux, a.cDocum, "
    s_Sentencia = s_Sentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(cDebeMN), 0), 2) AS nDebSol, "
    s_Sentencia = s_Sentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(cHaberMN), 0), 2) AS nHabSol, "
    s_Sentencia = s_Sentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(cDebeME), 0), 2) AS nDebDol, "
    s_Sentencia = s_Sentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(cHaberME), 0), 2) AS nHabDol "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpfiltros ", "")
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trptcctctahs a "
    s_Sentencia = s_Sentencia & "GROUP BY " & IIf((Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> ""), "a.codcco, ", "") & "a.codcta, a.CodAux, a.cDocum "
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "HAVING (ROUND(nDebSol - nHabSol, 2)" & s_Comparar & "0.00 OR ROUND(nDebDol - nHabDol, 2)" & s_Comparar & "0.00) "
    Else
      s_Sentencia = s_Sentencia & "HAVING (ROUND(ROUND(ISNULL(SUM(cDebeMN), 0), 2) - ROUND(ISNULL(SUM(cHaberMN), 0), 2), 2)" & s_Comparar & "0.00 "
      s_Sentencia = s_Sentencia & "OR ROUND(ROUND(ISNULL(SUM(cDebeME), 0), 2) - ROUND(ISNULL(SUM(cHaberME), 0), 2), 2)" & s_Comparar & "0.00) "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY " & IIf((Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> ""), "a.codcco, ", "") & "a.codcta, a.codaux, a.cdocum"
    pocnnMain.Execute s_Sentencia
  End If
   
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.* FROM " & ps_Prefijo & "trptcctctahs a"
    If cboInformacion.ListIndex <> 0 Then
      .Source = .Source & ", " & ps_Prefijo & "tmpfiltros b "
      .Source = .Source & "WHERE a.codcta=b.codcta "
      If (Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> "") Then
        .Source = .Source & "AND a.codcco=b.codcco "
      End If
      .Source = .Source & "AND a.codaux=b.codaux "
      .Source = .Source & "AND a.cdocum=b.cdocum"
    End If
    .Source = .Source & " ORDER BY " & IIf((Trim(txtDato(3).Text) <> "" And Trim(txtDato(4).Text) <> ""), "a.codcco, ", "") & "a.CodCta, a.CodAux, a.cDocum, a.MesPvs, a.FehOpe"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRCCtHst.rpt"
      .ParameterFields(1) = "ccquiebre;" & sQuiebre & ";true"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRCCtHst.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption, udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
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
  ' Elimino el archivo temporal
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#trptcctctahs_') DROP TABLE #trptcctctahs"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptcctctahs", s_Sentencia)
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpfiltros_') DROP TABLE #tmpfiltros"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpfiltros", s_Sentencia)
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
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2, 3, 4                        'Cambiar (añadir índices).
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
   Case 2                              'Cambiar (añadir índices).
    modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 3, 4                           'Cambiar (añadir índices).
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
   Case 2
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
   Case 3, 4
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstCoCCo
      .MoveFirst
      .Find "codcco='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
      End If
    End With
  End Select

End Function

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
  optTipoImpresion(0).Enabled = tbHabilitar
  optTipoImpresion(1).Enabled = tbHabilitar
  cmdImprimir(0).Enabled = tbHabilitar
  cmdImprimir(1).Enabled = tbHabilitar
  cmdConfig.Enabled = tbHabilitar
  cmdSalir.Enabled = tbHabilitar
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

