VERSION 5.00
Begin VB.Form frmRTp56Cab 
   Caption         =   "[título]"
   ClientHeight    =   2730
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFormato 
      Caption         =   "Formato "
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   45
      TabIndex        =   9
      Top             =   1320
      Width           =   2520
      Begin VB.OptionButton optPagina 
         Caption         =   "Horizontal"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1305
         TabIndex        =   11
         Top             =   315
         Width           =   1020
      End
      Begin VB.OptionButton optPagina 
         Caption         =   "Vertical"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1035
         TabIndex        =   13
         Top             =   315
         Width           =   1020
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango de Pagina "
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   0
      TabIndex        =   4
      Top             =   75
      Width           =   5055
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
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
         Left            =   1215
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
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
         Left            =   3540
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Fin : "
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Top             =   405
         Width           =   345
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Inicio : "
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   405
         Width           =   510
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
      ScaleWidth      =   5130
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2190
      Width           =   5130
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
         Picture         =   "frmrtp56cab.frx":0000
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
         Picture         =   "frmrtp56cab.frx":014A
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
         Picture         =   "frmrtp56cab.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRTp56Cab"
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
']

Private Sub Form_Load()
  On Error GoTo Err
  
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   
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
  
  
  '[Parámetros.                         'Cambiar.
  txtDato(0).MaxLength = 7
  txtDato(1).MaxLength = 7
  ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Beginning :", "End :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango Pagina", "Range Page")
  fraFormato.Caption = Choose(gsIdioma, "Formato Papel", "Paper Setup")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
   
  '[Datos predeterminados.              'Cambiar.
  txtDato(0) = 1: txtDato(1) = 1
  
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

Private Sub cmdImprimir_Click(Index As Integer)
  Dim nContador As Double
  Dim sSentencia As String, sRegistro As String
  Dim sReporte As String
  
  ppHabilitacion False
  ' Realizo las validaciones de los campos a actualizar
  If Not IsNumeric(txtDato(0).Text) Then Beep: MsgBox TEXT_8010, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If txtDato(0).Text <= 0 Then Beep: MsgBox Choose(gsIdioma, "Número de pagina inicial es invalido", "Number of initial pagina is not been worth"), vbExclamation: txtDato(0).SetFocus: Exit Sub
  If Not IsNumeric(txtDato(1).Text) Then Beep: MsgBox TEXT_8010, vbExclamation: txtDato(1).SetFocus: Exit Sub
  If txtDato(1).Text <= 0 Then Beep: MsgBox Choose(gsIdioma, "Número de pagina final es invalido", "Number of final page is not been worth"), vbExclamation: txtDato(1).SetFocus: Exit Sub
  If Not (CDec(txtDato(1).Text) >= CDec(txtDato(0).Text)) Then MsgBox Choose(gsIdioma, "Número de pagina final debe ser mayor o igual que pagina inicial", "Number of final page must be greater or just as page initial"), vbExclamation: txtDato(1).SetFocus: Exit Sub
  
  sReporte = IIf(optPagina(0).Value, "rptrcabver", "rptrcabhor")
  'Elimino la informacion del reporte
  sSentencia = "DELETE FROM cotmprpt "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND nomrpt='" & sReporte & "' "
  sSentencia = sSentencia & "AND usrcre='" & gsAbvUsr & "'"
  pocnnMain.Execute sSentencia
  For nContador = CDec(txtDato(0).Text) To CDec(txtDato(1).Text)
    sRegistro = Format(nContador, "0000000")
    sSentencia = "INSERT INTO cotmprpt (codemp, pdoano, codcta, nomrpt, usrcre) "
    sSentencia = sSentencia & "VALUES ('" & gsCodEmp & "', '" & gsAnoAct & "', '" & sRegistro & "', '" & sReporte & "', '" & gsAbvUsr & "')"
    pocnnMain.Execute sSentencia
  Next nContador
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT codemp, pdoano, codcta "
    .Source = .Source & "FROM cotmprpt "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND nomrpt='" & sReporte & "' "
    .Source = .Source & "AND usrcre='" & gsAbvUsr & "'"
    .Source = .Source & "ORDER BY codcta"
    .Open
   End With
   
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & sReporte & ".rpt"
      '[ Formulas adicionales
      sRegistro = FormatNumber(Left(Trim(gsRUCEmp), 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(6) = "mRucEmpresa='" & sRegistro & Mid(Trim(gsRUCEmp), 9) & "'"
      .Formulas(7) = "mDireccion='" & gsDirEmp & "'"
      .Formulas(8) = "mDistrito='" & gsLocEmp & "'"
      .Formulas(9) = "mActividad='" & gsGirEmp & "'"
      .Formulas(10) = "mRepresentante='" & gsRepEmp & "'"
      sRegistro = FormatNumber(Left(gsRepDNIEmp, 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(11) = "mDniRepresentante='" & sRegistro & Mid(gsRepDNIEmp, 9) & "'"
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
      .LoadReport gsRutRpt & "rptrcabhor.mrp"
      
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
  ppHabilitacion True
  'Elimino la informacion del reporte
  sSentencia = "DELETE FROM cotmprpt "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND nomrpt='" & sReporte & "' "
  sSentencia = sSentencia & "AND usrcre='" & gsAbvUsr & "'"
  pocnnMain.Execute sSentencia

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

Private Sub Form_Unload(Cancel As Integer)
   pocnnMain.Close
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  txtDato(Index).Text = IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)
  txtDato(Index).Text = FormatNumber(txtDato(Index).Text, 0)
End Sub

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
