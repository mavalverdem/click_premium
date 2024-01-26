VERSION 5.00
Begin VB.Form frmRCCtPdo 
   Caption         =   "[título]"
   ClientHeight    =   6090
   ClientLeft      =   2460
   ClientTop       =   1875
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7035
   Begin VB.Frame fraReferencia 
      Caption         =   " Referencia "
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   90
      TabIndex        =   18
      Top             =   2865
      Width           =   4215
      Begin VB.TextBox txtReferencia 
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
         Left            =   855
         MaxLength       =   20
         TabIndex        =   20
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Pedido :"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   19
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   90
      TabIndex        =   32
      Top             =   4770
      Width           =   3615
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   33
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
         TabIndex        =   34
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Historico"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   2325
         TabIndex        =   35
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.CheckBox chkRango 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1485
      TabIndex        =   25
      Top             =   3660
      Width           =   180
   End
   Begin VB.Frame fraRngPeriodo 
      Caption         =   " Rango Periodos "
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   90
      TabIndex        =   24
      Top             =   3645
      Width           =   4215
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   27
         Text            =   "Año Inicio"
         Top             =   300
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   855
         TabIndex        =   30
         Text            =   "Año Final"
         Top             =   645
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   28
         Text            =   "Mes Inicio"
         Top             =   300
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   31
         Text            =   "Mes Final"
         Top             =   645
         Width           =   1710
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   29
         Top             =   690
         Width           =   765
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   36
      Top             =   4770
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
         Left            =   1020
         TabIndex        =   38
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5745
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2925
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H00800000&
      Height          =   2805
      Left            =   0
      TabIndex        =   4
      Top             =   60
      Width           =   6990
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   4
         Left            =   135
         TabIndex        =   16
         Top             =   2385
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   4
         Left            =   6585
         Picture         =   "frmRCCtPdo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   13
         Top             =   1755
         Width           =   630
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1395
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   6585
         Picture         =   "frmRCCtPdo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1410
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   6585
         Picture         =   "frmRCCtPdo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1755
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRCCtPdo.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   495
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   495
         Width           =   945
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   825
         Width           =   945
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRCCtPdo.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   825
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
         Height          =   285
         Index           =   4
         Left            =   1380
         TabIndex        =   17
         Top             =   2385
         Width           =   5205
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Auxiliar"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   2160
         Width           =   495
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
         Height          =   285
         Index           =   3
         Left            =   750
         TabIndex        =   14
         Top             =   1755
         Width           =   5835
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
         Height          =   285
         Index           =   2
         Left            =   750
         TabIndex        =   12
         Top             =   1410
         Width           =   5835
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
         Height          =   285
         Index           =   0
         Left            =   1065
         TabIndex        =   7
         Top             =   495
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
         Height          =   285
         Index           =   1
         Left            =   1065
         TabIndex        =   9
         Top             =   825
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
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centros de Costo"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   1170
         Width           =   1215
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
      ScaleWidth      =   7035
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5550
      Width           =   7035
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
         Picture         =   "frmRCCtPdo.frx":0852
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
         Picture         =   "frmRCCtPdo.frx":099C
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
         Picture         =   "frmRCCtPdo.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   4980
      TabIndex        =   21
      Top             =   2970
      Width           =   675
   End
End
Attribute VB_Name = "frmRCCtPdo"
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
 ']

 '[Parámetros.                         'Cambiar.
   With TxtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      For dnContador = 2 To 3
         .Item(dnContador).DataField = "CodCCo"
         .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      .Item(4).DataField = "CodAux"
      .Item(4).MaxLength = porstTGAux.Fields(.Item(4).DataField).DefinedSize
   End With
 ']
   
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(7, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Centro de Costo :", "Auxiliar :", "Pedido :", "Moneda :", "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Cost Center :", "Auxiliary :", "Order :", "Currency :", "Beginning :", "End :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraReferencia.Caption = Choose(gsIdioma, " Referencia ", " Reference ")
  fraRngPeriodo.Caption = Choose(gsIdioma, "Rango Periodos", "Range of Periods")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  OptTipo(2).Caption = Choose(gsIdioma, "Histórico", "Historical")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
 
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      TxtDato(1).Text = !CodCta
      .MoveFirst
      TxtDato(0).Text = !CodCta
   End With
   With porstCoCCo
      .MoveLast
      TxtDato(3).Text = !codcco
      .MoveFirst
      TxtDato(2).Text = !codcco
   End With
   
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If TxtDato(0).Text <> "" Then ppAyuDet 0
   If TxtDato(1).Text <> "" Then ppAyuDet 1
   If TxtDato(2).Text <> "" Then ppAyuDet 2
   If TxtDato(3).Text <> "" Then ppAyuDet 3
   If TxtDato(4).Text <> "" Then ppAyuDet 4
  
  'Otros.
   OptTipo(0).Value = True
   
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
   porstCoCCo.Close
   porstCOCta.Close
   porstTGAux.Close
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
      TxtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim dnContador As Integer, s_Sql As String
  Dim s_Sentencia As String, s_Moneda As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
       
  s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
  s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
    s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End fiscal year must be equal or more than Opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
    If (s_AnoFin = s_AnoIni) And Not (s_Mes <= s_Ano) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  End If
  ppHabilitacion False
    
  ' Elimino y genero el archivo temporal del reporte
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpRptCtePdo_') DROP TABLE #tmpRptCtePdo"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRptCtePdo", s_Sentencia)
  s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE tmpRptCtePdo ", "")
  If OptTipo(2).Value Then
    s_Sentencia = s_Sentencia & "SELECT a.codemp, a.codaux, a.pdocpr, a.feedoc, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.abvtdc,'-',a.serdoc,'-',a.nrodoc)", "(b.abvtdc+'-'+a.serdoc+'-'+a.nrodoc)") & " AS cDocumcpr, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "a.gloite", "a.gloitex") & " AS glodoc, a.codcta, a.codcco, "
    s_Sentencia = s_Sentencia & "(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
    s_Sentencia = s_Sentencia & "ROUND((a.impmn)*(CASE a.tpoctb WHEN '" & TPOCTB_HAB & "' THEN -1 ELSE 1 END), 2) AS impcpr_mn, "
    s_Sentencia = s_Sentencia & "ROUND((a.impme)*(CASE a.tpoctb WHEN '" & TPOCTB_HAB & "' THEN -1 ELSE 1 END), 2) AS impcpr_me "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpRptCtePdo ")
    s_Sentencia = s_Sentencia & "FROM cocpbdet a "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCCo=d.CodCCo "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux e ON a.codemp=e.codemp AND a.codaux=e.codaux "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    ' Inicializo la informacion rango periodos
    s_AnoIni = Format(Trim(Val(gsAnoAct) - 1), "0000")
    s_AnoFin = gsAnoAct
    s_Mes = gsMesAct
    s_Ano = gsMesAct
    If chkRango.Value = vbChecked Then
      s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
      s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
      s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
      s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & ">='" & s_AnoIni & s_Mes & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & s_AnoFin & s_Ano & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.pdocpr, '') <>'' "
    s_Sentencia = s_Sentencia & "AND a.codcta BETWEEN '" & TxtDato(0).Text & "' AND '" & TxtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND a.codcco BETWEEN '" & TxtDato(2).Text & "' AND '" & TxtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codcta, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "", "") & "mid(a.codcta,1,1)<>'4' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codcco, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codaux, '')<>'' "
    If Trim(TxtDato(4).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.codaux='" & TxtDato(4).Text & "' "
    End If
    If Trim(txtReferencia.Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.pdocpr='" & txtReferencia.Text & "' "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY a.codaux, a.pdocpr"
  Else
    s_Sentencia = s_Sentencia & "SELECT a.codemp, a.codaux, a.pdocpr, a.codcta, a.codcco, "
    s_Sentencia = s_Sentencia & "ROUND(SUM((" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.impmn, 0)*(CASE a.tpoctb WHEN '" & TPOCTB_HAB & "' THEN -1 ELSE 1 END))), 2) AS impcpr_mn, "
    s_Sentencia = s_Sentencia & "ROUND(SUM((" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.impme, 0)*(CASE a.tpoctb WHEN '" & TPOCTB_HAB & "' THEN -1 ELSE 1 END))), 2) AS impcpr_me "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpRptCtePdo ")
    s_Sentencia = s_Sentencia & "FROM cocpbdet a "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCCo=d.CodCCo "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux e ON a.codemp=e.codemp AND a.codaux=e.codaux "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    ' Inicializo la informacion rango periodos
    s_AnoIni = Format(Trim(Val(gsAnoAct) - 1), "0000")
    s_AnoFin = gsAnoAct
    s_Mes = gsMesAct
    s_Ano = gsMesAct
    If chkRango.Value = vbChecked Then
      s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
      s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
      s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
      s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & ">='" & s_AnoIni & s_Mes & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & s_AnoFin & s_Ano & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.pdocpr, '') <>'' "
    s_Sentencia = s_Sentencia & "AND a.codcta BETWEEN '" & TxtDato(0).Text & "' AND '" & TxtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND a.codcco BETWEEN '" & TxtDato(2).Text & "' AND '" & TxtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codcta, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "", "") & "mid(a.codcta,1,1)<>'4' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codcco, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codaux, '')<>'' "
    If Trim(TxtDato(4).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.codaux='" & TxtDato(4).Text & "' "
    End If
    If Trim(txtReferencia.Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.pdocpr='" & txtReferencia.Text & "' "
    End If
    s_Sentencia = s_Sentencia & "GROUP BY a.codemp, a.codaux, a.pdocpr, a.codcta, a.codcco "
    s_Sentencia = s_Sentencia & "ORDER BY a.codaux, a.pdocpr"
  End If
  pocnnMain.Execute s_Sentencia
  
  
  If OptTipo(0).Value Then
    s_Sentencia = "SELECT a.codaux, e.razaux, concat(a.coddpe,a.pdocpr) as pdocpr, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo, "
    s_Sentencia = s_Sentencia & "a.fehpdo, x.codcta, x.codcco, "
    s_Sentencia = s_Sentencia & "(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS ctpomon, a.impdife, "
'ini 2014-07-22 error otros rpt
'    s_Sentencia = s_Sentencia & "ROUND(x.impcta_mn-" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_mn, 0), 2) AS saldomn, "
'    s_Sentencia = s_Sentencia & "ROUND(x.impcta_me-" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_me, 0), 2) AS saldome, "
    s_Sentencia = s_Sentencia & "ROUND((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (x.impcta_mn*(" & gnPctIGV & "/100))+x.impcta_mn ELSE x.impcta_mn END)-" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_mn, 0), 2) AS saldomn, "
    s_Sentencia = s_Sentencia & "ROUND((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (x.impcta_me*(" & gnPctIGV & "/100))+x.impcta_me ELSE x.impcta_me END)-" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_me, 0), 2) AS saldome, "
'fin 2014-07-22 error otros rpt
    s_Sentencia = s_Sentencia & "a.nrointerno "
    s_Sentencia = s_Sentencia & "FROM copdocpr a "
    s_Sentencia = s_Sentencia & "LEFT JOIN copdocprcta x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and concat(a.coddpe,a.pdocpr)=concat(x.coddpe,x.pdocpr) "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & ps_Prefijo & "tmpRptCtePdo b ON a.codemp=b.codemp AND a.codaux=b.codaux AND concat(a.coddpe,a.pdocpr)=b.pdocpr and x.codcco=b.codcco and x.codcta=b.codcta "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND x.CodCta=c.CodCta "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND x.CodCCo=d.CodCCo "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux e ON a.codemp=e.codemp AND a.codaux=e.codaux "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    ' Inicializo la informacion rango periodos
    s_AnoIni = Format(Trim(Val(gsAnoAct) - 1), "0000")
    s_AnoFin = gsAnoAct
    s_Mes = gsMesAct
    s_Ano = gsMesAct
    If chkRango.Value = vbChecked Then
      s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
      s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
      s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
      s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & ">='" & s_AnoIni & s_Mes & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & s_AnoFin & s_Ano & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.pdocpr, '') <>'' "
    s_Sentencia = s_Sentencia & "AND x.codcta BETWEEN '" & TxtDato(0).Text & "' AND '" & TxtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND x.codcco BETWEEN '" & TxtDato(2).Text & "' AND '" & TxtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.codcta, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.codcco, '')<>'' "
    If Trim(TxtDato(4).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.codaux='" & TxtDato(4).Text & "' "
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codaux, '')<>'' "
    End If
    If Trim(txtReferencia.Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.pdocpr='" & txtReferencia.Text & "' "
    End If
    'ini 2014-07-21 adicion igv gravado
'    If ps_Plataforma = pSrvMySql Then
'      s_Sentencia = s_Sentencia & "AND ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN (a.impmn-IFNULL(b.impcpr_mn, 0)) ELSE (a.impme-IFNULL(b.impcpr_me, 0)) END, 2)>a.impdife "
'    ElseIf ps_Plataforma = pSrvSql Then
'      s_Sentencia = s_Sentencia & "AND ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN (a.impmn-ISNULL(b.impcpr_mn, 0)) ELSE (a.impme-ISNULL(b.impcpr_me, 0)) END, 2)>a.impdife "
'    End If
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "AND ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "' "
      s_Sentencia = s_Sentencia & "THEN ((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (a.impmn*(" & gnPctIGV & "/100))+a.impmn ELSE a.impmn END)-IFNULL(b.impcpr_mn, 0))  "
      s_Sentencia = s_Sentencia & "ELSE ((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (a.impme*(" & gnPctIGV & "/100))+a.impme ELSE a.impme END)-IFNULL(b.impcpr_me, 0)) END, 2)>a.impdife "
    ElseIf ps_Plataforma = pSrvSql Then
      s_Sentencia = s_Sentencia & "AND ROUND(CASE a.tpomon WHEN '" & TPOMON_NAC & "'  "
      s_Sentencia = s_Sentencia & "THEN ((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (a.impmn*(" & gnPctIGV & "/100))+a.impmn ELSE a.impmn END)-ISNULL(b.impcpr_mn, 0)) "
      s_Sentencia = s_Sentencia & "ELSE ((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (a.impme*(" & gnPctIGV & "/100))+a.impme ELSE a.impme END)-ISNULL(b.impcpr_me, 0)) END, 2)>a.impdife "
    End If
    'fin 2014-07-21 adicion igv gravado
    s_Sentencia = s_Sentencia & "ORDER BY a.codaux, a.pdocpr"
  ElseIf OptTipo(1).Value Then
    s_Sentencia = "SELECT a.codaux, e.razaux, "
'ini 2014-07-22 error otros rpt
'    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.impcta_mn, 0)), 2) AS impcta_mn, "
'    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_mn, 0)), 2) AS impcpr_mn, "
'    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.impcta_me, 0)), 2) AS impcta_me, "
'    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_me, 0)), 2) AS impcpr_me "
    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (x.impcta_mn*(" & gnPctIGV & "/100))+x.impcta_mn ELSE x.impcta_mn END), 0)), 2) AS impcta_mn, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_mn, 0)), 2) AS impcpr_mn, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (x.impcta_me*(" & gnPctIGV & "/100))+x.impcta_me ELSE x.impcta_me END), 0)), 2) AS impcta_me, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.impcpr_me, 0)), 2) AS impcpr_me "
'fin 2014-07-22 error otros rpt
    s_Sentencia = s_Sentencia & "FROM copdocpr a "
    s_Sentencia = s_Sentencia & "LEFT JOIN copdocprcta x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and concat(a.coddpe,a.pdocpr)=concat(x.coddpe,x.pdocpr) "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & ps_Prefijo & "tmpRptCtePdo b ON a.codemp=b.codemp AND a.codaux=b.codaux AND concat(a.coddpe,a.pdocpr)=b.pdocpr and x.codcco=b.codcco and x.codcta=b.codcta "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON x.codemp=c.codemp AND x.pdoano=c.pdoano AND x.CodCta=c.CodCta "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo d ON x.codemp=d.codemp AND x.pdoano=d.pdoano AND x.CodCCo=d.CodCCo "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux e ON a.codemp=e.codemp AND a.codaux=e.codaux "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    ' Inicializo la informacion rango periodos
    s_AnoIni = Format(Trim(Val(gsAnoAct) - 1), "0000")
    s_AnoFin = gsAnoAct
    s_Mes = gsMesAct
    s_Ano = gsMesAct
    If chkRango.Value = vbChecked Then
      s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
      s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
      s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
      s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & s_AnoFin & s_Ano & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.pdocpr, '') <>'' "
    s_Sentencia = s_Sentencia & "AND x.codcta BETWEEN '" & TxtDato(0).Text & "' AND '" & TxtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND x.codcco BETWEEN '" & TxtDato(2).Text & "' AND '" & TxtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.codcta, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.codcco, '')<>'' "
    If Trim(TxtDato(4).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.codaux='" & TxtDato(4).Text & "' "
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codaux, '')<>'' "
    End If
    If Trim(txtReferencia.Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.pdocpr='" & txtReferencia.Text & "' "
    End If
    s_Sentencia = s_Sentencia & "GROUP BY a.CodAux, e.RazAux "
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "HAVING (ROUND(impcta_mn-impcpr_mn, 2) <> 0.00 OR ROUND(impcta_me-impcpr_me, 2) <> 0.00) "
    ElseIf ps_Plataforma = pSrvSql Then
      s_Sentencia = s_Sentencia & "HAVING (ROUND(ROUND(SUM(ISNULL(x.impcta_mn, 0)), 2) - ROUND(SUM(ISNULL(b.impcpr_mn, 0)), 2), 2) <> 0.00) "
      s_Sentencia = s_Sentencia & "OR (ROUND(ROUND(SUM(ISNULL(x.impcta_me, 0)), 2) - ROUND(SUM(ISNULL(b.impcpr_me, 0)), 2), 2) <> 0.00) "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY a.codaux, e.razaux"
  ElseIf OptTipo(2).Value Then
    s_Sentencia = "SELECT a.codaux, e.razaux, concat(a.coddpe,a.pdocpr) as pdocpr, a.fehpdo AS fecha, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS glosa, x.codcta, x.codcco, space(20) AS cdocumcpr, "
    s_Sentencia = s_Sentencia & "(CASE a.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cmoneda, "
    'ini 2014-07-21 adicion igv gravado
    's_Sentencia = s_Sentencia & "x.impcta_mn AS imppdomn, ROUND(0, 2) AS impcprmn, "
    's_Sentencia = s_Sentencia & "x.impcta_me AS imppdome, ROUND(0, 2) AS impcprme, '0' AS orden, a.nrointerno "
    s_Sentencia = s_Sentencia & "(CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (x.impcta_mn*(" & gnPctIGV & "/100))+x.impcta_mn ELSE x.impcta_mn END) AS imppdomn, "
    s_Sentencia = s_Sentencia & "ROUND(0, 2) AS impcprmn, "
    s_Sentencia = s_Sentencia & "(CASE a.tpoigv WHEN '" & CODPDO_IGVG & "' THEN (x.impcta_me*(" & gnPctIGV & "/100))+x.impcta_me ELSE x.impcta_me END) AS imppdome, "
    s_Sentencia = s_Sentencia & "ROUND(0, 2) AS impcprme,  "
    s_Sentencia = s_Sentencia & " '0' AS orden, a.nrointerno "
    'fin 2014-07-21 adicion igv gravado
    s_Sentencia = s_Sentencia & "FROM copdocpr a "
    s_Sentencia = s_Sentencia & "LEFT JOIN copdocprcta x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and concat(a.coddpe,a.pdocpr)=concat(x.coddpe,x.pdocpr) "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND x.CodCta=c.CodCta "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCCo d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND x.CodCCo=d.CodCCo "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux e ON a.codemp=e.codemp AND a.codaux=e.codaux "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    ' Inicializo la informacion rango periodos
    s_AnoIni = Format(Trim(Val(gsAnoAct) - 1), "0000")
    s_AnoFin = gsAnoAct
    s_Mes = gsMesAct
    s_Ano = gsMesAct
    If chkRango.Value = vbChecked Then
      s_AnoIni = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(0), gsAnoAct), 4)
      s_AnoFin = Right(IIf(chkRango.Value = vbChecked, cmbPeriodo(1), gsAnoAct), 4)
      s_Mes = Format(cmbPeriodo(2).ListIndex, "00")
      s_Ano = Format(cmbPeriodo(3).ListIndex, "00")
    End If
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & ">='" & s_AnoIni & s_Mes & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & s_AnoFin & s_Ano & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.pdocpr, '') <>'' "
    s_Sentencia = s_Sentencia & "AND x.codcta BETWEEN '" & TxtDato(0).Text & "' AND '" & TxtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND x.codcco BETWEEN '" & TxtDato(2).Text & "' AND '" & TxtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.codcta, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "", "") & "mid(x.codcta,1,1)<>'4' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(x.codcco, '')<>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codaux, '')<>'' "
    If Trim(TxtDato(4).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.codaux='" & TxtDato(4).Text & "' "
    End If
    If Trim(txtReferencia.Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.pdocpr='" & txtReferencia.Text & "' "
    End If
    s_Sentencia = s_Sentencia & "UNION "
    s_Sentencia = s_Sentencia & "SELECT a.codaux, e.razaux, a.pdocpr, a.feedoc AS fecha, a.glodoc AS glosa, a.codcta, a.codcco, cdocumcpr AS cdocumcpr, "
    s_Sentencia = s_Sentencia & "a.csigno AS cmoneda, "
    s_Sentencia = s_Sentencia & "ROUND(0, 2) AS imppdomn, a.impcpr_mn AS impcprmn, "
    s_Sentencia = s_Sentencia & "ROUND(0, 2) AS imppdome, a.impcpr_me AS impcprme, '1' AS orden, b.nrointerno "
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "tmprptctepdo a "
    s_Sentencia = s_Sentencia & "INNER JOIN copdocpr b ON a.codaux=b.codaux AND a.pdocpr=concat(b.coddpe,b.pdocpr) AND a.codemp=b.codemp "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.pdoano, b.mespvs)", "(b.pdoano+b.mespvs)") & ">='" & s_AnoIni & s_Mes & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.pdoano, b.mespvs)", "(b.pdoano+b.mespvs)") & "<='" & s_AnoFin & s_Ano & "' "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux e ON a.codemp=e.codemp AND a.codaux=e.codaux "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.pdocpr, '') <>'' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.codaux, '')<>'' "
    If Trim(TxtDato(4).Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.codaux='" & TxtDato(4).Text & "' "
    End If
    If Trim(txtReferencia.Text) <> "" Then
      s_Sentencia = s_Sentencia & "AND a.pdocpr='" & txtReferencia.Text & "' "
    End If
    s_Sentencia = s_Sentencia & "ORDER BY codaux, pdocpr, fecha, orden"
  End If
  ' Recordset del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = s_Sentencia
    .Open
  End With
  
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(OptTipo(0).Value, OptTipo(0).Caption, IIf(OptTipo(1).Value, OptTipo(1).Caption, OptTipo(2).Caption)) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptrcctpdodet.rpt", IIf(OptTipo(1).Value, "rptrcctpdores.rpt", "rptrcctpdohst.rpt"))
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  End If
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpRptCtePdo_') DROP TABLE #tmpRptCtePdo"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRptCtePdo", s_Sentencia)
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
   TxtDato(Index).SelStart = 0
   TxtDato(Index).SelLength = TxtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(TxtDato(Index))) + 1 = TxtDato(Index).MaxLength Then
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
   Case 0, 1, 2, 3, 4                        'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
    modAyuBus.Cta_Cod "", TxtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + fraRangos.Left + TxtDato(tnIndex).Left
    TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2, 3                           'Cambiar (añadir índices).
    modAyuBus.CCo_Cod "", TxtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + fraRangos.Left + TxtDato(tnIndex).Left
    TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 4                              'Cambiar (añadir índices).
    modAyuBus.Aux_Det "", TxtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + fraRangos.Left + TxtDato(tnIndex).Left
    TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 0, 1
    If TxtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstCOCta
      .MoveFirst
      .Find "CodCta='" & TxtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
      End If
    End With
   Case 2, 3
    If TxtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstCoCCo
      .MoveFirst
      .Find "CodCCo='" & TxtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
      End If
    End With
   Case 4
    If TxtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstTGAux
      .MoveFirst
      .Find "CodAux='" & TxtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!razAux), "", !razAux)
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
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

Private Sub txtReferencia_GotFocus()
  txtReferencia.SelStart = 0
  txtReferencia.SelLength = txtReferencia.MaxLength
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
  '[ARREGLAR: Retrocede si Shift está presionado.
  If Len(Trim(txtReferencia)) + 1 = txtReferencia.MaxLength Then
    SendKeys "{TAB}"
  End If
  ']ARREGLAR.
End Sub

