VERSION 5.00
Begin VB.Form frmLAux 
   Caption         =   "[título]"
   ClientHeight    =   2850
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimirCuenta 
      Caption         =   "&Cuenta Bancos"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Picture         =   "frmLAux.frx":0000
      TabIndex        =   16
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1035
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   315
         Width           =   915
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6870
         Picture         =   "frmLAux.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6870
         Picture         =   "frmLAux.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Auxiliares"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   660
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
         TabIndex        =   11
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
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   10
         Top             =   855
         Width           =   5520
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
      ScaleWidth      =   7320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2310
      Width           =   7320
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Excel"
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
         Index           =   2
         Left            =   3720
         Picture         =   "frmLAux.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   6120
         Picture         =   "frmLAux.frx":0C76
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
         Picture         =   "frmLAux.frx":0DC0
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
         Picture         =   "frmLAux.frx":12F2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmLAux"
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
']

Private Sub cmdImprimirCuenta_Click()
ppHabilitacion False
    
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT TGAux.CodAux as d1,razaux as d2,tpoper as d3,indcli as d4,indprv as d5,indotr as d6,nroctacte as d7,nrocci as d8,detbco as d9,tpocta as d10,tpomon as d11,rucaux as d12 "
    .Source = .Source & " FROM TGAux "
    .Source = .Source & " LEFT JOIN coctaban ON TGaux.codemp=coctaban.codemp AND TGAux.CodAux = coctaban.CodAux "
    .Source = .Source & " LEFT JOIN cobco ON coctaban.codemp=cobco.codemp AND coctaban.Codbco = cobco.codbco "
    .Source = .Source & " WHERE TGAux.codemp='" & gsCodEmp & "' "
    .Source = .Source & " AND TGAux.CodAux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & " ORDER BY 2 asc"
    .Open
  End With
  
  'usDEstino = PRN_DEST_GRAF
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLAuxcta.rpt"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
    End With
  ppHabilitacion True

End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
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
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodAux"
         .Item(dnContador).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Auxiliares")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Auxiliarys")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstTGAux
      .MoveLast
      txtDato(1).Text = !codaux
      .MoveFirst
      txtDato(0).Text = !codaux
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
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
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub
'ini 2016-07-22 excel auxiliar
Private Sub cmdImprimir_Click(Index As Integer)
fcmdImprimir (Index)
End Sub

Private Sub fcmdImprimir(Index As Integer)
  ppHabilitacion False
    
  With porstMRp
'    If Index <> 2 Then
'    If .State = adStateOpen Then .Close
'    Else
    If .State = adStateOpen Then
      .Close
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
      
    End If
'    End If
    .Source = "SELECT TGAux.CodAux, TGAux.RazAux, TGAux.RucAux, "
    .Source = .Source & "(Case TGAux.TpoPer when 'J' then 'Juridica' when 'N' then 'Natural' End) AS cTpoPer, "
    .Source = .Source & "TGAuxNat.NomAux, TGAuxNat.ApePatAux, TGAuxNat.ApeMatAux, "
    .Source = .Source & "(Case TGAux.IndCli  when '1' then 'X'  End) as cIndCli, "
    .Source = .Source & "(Case TGAux.IndPrv  when '1' then 'X' End) as cIndPrv, "
    .Source = .Source & "(Case TGAux.IndOtr  when '1' then 'X' End) as cIndOtr, "
    .Source = .Source & "(Case TGAux.EstAux when 'A' Then 'Activo' Else 'Inactivo' End) as vEstAux "
'ini 2016-07-22 excel auxiliar
    .Source = .Source & ",diraux "
'fin 2016-07-22 excel auxiliar
    .Source = .Source & "FROM TGAux "
    .Source = .Source & "LEFT JOIN TGAuxNat ON TGaux.codemp=TGAuxNat.codemp AND TGAux.CodAux = TGAuxNat.CodAux "
    .Source = .Source & "WHERE TGAux.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND TGAux.CodAux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "ORDER BY TGAux.CodAux"
    .Open
  End With
  If Index = 2 Then
    pExporta
  Else
    usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
    If usDEstino = PRN_DEST_GRAF Then
      gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
      With frmMain.rptMain
        '[Datos y parámetros del reporte.  'Cambiar.
        .ReportFileName = gsRutRpt & "rptLAux.rpt"
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
        .LoadReport gsRutRpt & "rptLAux.mrp"
        
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
  End If
  ppHabilitacion True

End Sub
Private Sub pExporta()
    'Dim xArchPeriodo As String
    'xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet

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
        .Cells(nRowI, 1).Value = "Listado de Auxiliares"
        nRowI = nRowI + 2
        Dim x1 As Integer
        .Cells(nRowI, 1).Value = "Código"
        .Cells(nRowI, 2).Value = "Razón Social"
        '.Cells(nRowI, 3).Value = "RUC"
        .Cells(nRowI, 4 - 1).Value = "RUC"
        .Cells(nRowI, 5 - 1).Value = "Tipo de Persona"
        .Cells(nRowI, 6 - 1).Value = "Nombres"
        .Cells(nRowI, 7 - 1).Value = "Apellido Paterno"
        .Cells(nRowI, 8 - 1).Value = "Apellido Materno"
        .Cells(nRowI, 9 - 1).Value = "Cliente"
        .Cells(nRowI, 10 - 1).Value = "Proveedor"
        .Cells(nRowI, 11 - 1).Value = "Otros"
        .Cells(nRowI, 12 - 1).Value = "Estado"
        .Cells(nRowI, 13 - 1).Value = "Direccion"
        
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstMRp
        .Columns.AutoFit ' ajusta el ancho de las columnas
      
    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing


'fin exporta datos a excel

'   porstTmp.Close
'   pocnnTmp.Close
'   Set porstTmp = Nothing
'   Set pocnnTmp = Nothing

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
'fin 2016-07-22 excel auxiliar

Private Sub cmdImprimir_Click_2016_07_22(Index As Integer)
  ppHabilitacion False
    
  With porstMRp
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
    .Source = .Source & "WHERE TGAux.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND TGAux.CodAux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "ORDER BY TGAux.CodAux"
    .Open
  End With
  
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLAux.rpt"
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
      .LoadReport gsRutRpt & "rptLAux.mrp"
      
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
'   Select Case Index    'Completa con ceros a la izquierda.
'   Case 0, 1                           'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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

