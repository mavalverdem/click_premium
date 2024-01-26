VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRCerIR4 
   Caption         =   "[título]"
   ClientHeight    =   2730
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCertificado 
      Caption         =   "Otro Formato"
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
      Picture         =   "frmRCerIR4.frx":0000
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame fraFechaRep 
      Caption         =   "Fecha de Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2760
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   66125825
         CurrentDate     =   38385
      End
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   9
      Top             =   1440
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   840
      Left            =   0
      TabIndex        =   6
      Top             =   135
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6885
         Picture         =   "frmRCerIR4.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   325
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
         Top             =   315
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
         TabIndex        =   8
         Top             =   315
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
      ScaleWidth      =   7290
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2190
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
         Picture         =   "frmRCerIR4.frx":06DC
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
         Picture         =   "frmRCerIR4.frx":0826
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
         Picture         =   "frmRCerIR4.frx":0D58
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRCerIR4"
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
Private porsTmpRp As ADODB.Recordset
']

Private Sub cmdCertificado_Click()
Dim sSentencia  As String, sDireccion As String
  Dim sRepresentante As String, sDocumento As String
  Dim porsTemporal As New ADODB.Recordset

  ppHabilitacion False
  
  ' Obtengo los datos de la empresa
  
  sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(direccion, '') AS direccion, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(repapepaterno, ''), ' ', IFNULL(repapematerno, ''), ' ', IFNULL(repnombre, ''))", "(ISNULL(repapepaterno, '')+' '+ISNULL(repapematerno, '')+' '+ISNULL(repnombre, ''))") & " AS representante, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(repdocumento, '') AS documento "
  sSentencia = sSentencia & "FROM tgemp "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "'"
  With porsTemporal
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = sSentencia
    .Open , CONNSTRG & gsNomBDC
    sDireccion = !direccion
    sRepresentante = !representante
    sDocumento = !documento
    .Close
  End With
  
  ' Genero la información
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptceri4ta", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptceri4ta') DROP TABLE #trptceri4ta")
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE #") & "trptceri4ta ("
  'sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS ", "CREATE TABLE #") & "trptceri4ta ("
  sSentencia = sSentencia & "CodAux varchar(11) NOT NULL default '', "
  sSentencia = sSentencia & "RazAux varchar(80) default NULL, "
  sSentencia = sSentencia & "diraux varchar(80) default NULL, "
  sSentencia = sSentencia & "RucAux varchar(11) default NULL, "
  sSentencia = sSentencia & "rubro varchar(40) default NULL, "
  sSentencia = sSentencia & "moneda varchar(3) default NULL, "
  sSentencia = sSentencia & "ImpBru_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpNet_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpIR4_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpIES_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpORt_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "importepen decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "rentapen decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "cLetras_1 varchar(80) NOT NULL default '', "
  sSentencia = sSentencia & "cLetras_2 varchar(80) NOT NULL default '',"
  sSentencia = sSentencia & "importesuje decimal(12,2) NULL default '0.00')"
  pocnnMain.Execute sSentencia
  ' Seleciono los registros
  sSentencia = "INSERT INTO " & ps_Prefijo & "trptceri4ta "
  sSentencia = sSentencia & "SELECT a.CodAux, b.RazAux, b.diraux, b.RucAux, b.rubro, 'S/.' AS moneda, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpBru_MN), 0) as ImpBru_MN, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpNet_MN), 0) as ImpNet_MN, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpIR4_MN), 0) as ImpIR4_MN, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpIES_MN), 0) as ImpIES_MN, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpORt_MN), 0) as ImpORt_MN, 0.00 AS importepen, 0.00 AS rentapen, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "REPEAT", "REPLICATE") & "(' ', 80) AS cLetras_1, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "REPEAT", "REPLICATE") & "(' ', 80) AS cLetras_2, "
  sSentencia = sSentencia & "(select sum(x.impbru_mn) from cohprdoc x where x.codemp='" & gsCodEmp & "' AND x.pdoano='" & gsAnoAct & "' and x.impir4_mn>0 and x.codaux=a.codaux) as importesuje "
  sSentencia = sSentencia & "FROM (cohprdoc a "
  sSentencia = sSentencia & "LEFT JOIN tgaux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  If Trim(txtDato(0).Text) <> "" Then
     sSentencia = sSentencia & "AND a.CodAux = '" & Trim(txtDato(0).Text) & "' "
  End If
  sSentencia = sSentencia & "GROUP BY a.CodAux, b.RazAux, b.diraux, b.RucAux, b.rubro "
  sSentencia = sSentencia & "ORDER BY a.CodAux"
  pocnnMain.Execute sSentencia
  'Actualizo Ultima Columna si es nulo
  sSentencia = "update " & ps_Prefijo & "trptceri4ta "
  sSentencia = sSentencia & "set importesuje=0 where importesuje is null "
  pocnnMain.Execute sSentencia
  ' Ingreso los datos al temporal
  ActualizaTemporal
  ' Obtengo la información
  With porstMRp
    If .State = adStateOpen Then .Close
     .Source = "SELECT * "
     .Source = .Source & "FROM " & ps_Prefijo & "trptceri4ta "
     .Source = .Source & "ORDER BY CodAux"
    .Open
  End With

    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRCerIR42010.rpt"
      'Fórmular propias.
      .Formulas(6) = "pEjercicio ='EJERCICIO " & gsAnoAct & "'"
      .Formulas(7) = "pRucEmpresa = '" & gsRUCEmp & "'"
      .Formulas(8) = "pDireccion=' " & sDireccion & "'"
      .Formulas(9) = "Representante='" & sRepresentante & "'"
      .Formulas(10) = "RepDocumento='" & sDocumento & "'"
      .Formulas(11) = "pFecha = '" & DTPfecha & "'"
      .Formulas(12) = "nFecha = '" & gsLocEmp & ", " & Day(DTPfecha) & " de " & Format(DTPfecha, " mmmm ") & " del " & Year(DTPfecha) & "'"
      '.Formulas(8) = "pFecha = 'LIMA " & Day(Date) & " DE " & UCase(Format(Date, " mmmm ")) & " DEL " & Year(Date) & "'"
      ']
      
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = crptToWindow
      .Action = 1
      
    End With
  ' Elimino la tabla temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptceri4ta", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptceri4ta') DROP TABLE #trptceri4ta")

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
      For dnContador = 0 To 0
         .Item(dnContador).DataField = "CodAux"
         .Item(dnContador).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  fraAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  fraFechaRep.Caption = Choose(gsIdioma, "Fecha de Impresión", "Printing Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']

' [Datos predeterminados.              'Cambiar.
' Límites de rangos.
'   With porstTGAux
'      .MoveLast
'      txtDato(1).Text = !CodAux
'      .MoveFirst
'      txtDato(0).Text = !CodAux
'   End With
' Busca detalle de códigos            '(habilitar/deshabilitar).

   If txtDato(0).Text <> "" Then ppAyuDet 0
  
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
   DTPfecha.Value = Date
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
   Case 0
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sSentencia  As String, sDireccion As String
  Dim sRepresentante As String, sDocumento As String
  Dim porsTemporal As New ADODB.Recordset

  ppHabilitacion False
  
  ' Obtengo los datos de la empresa
  sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(direccion, '') AS direccion, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(repapepaterno, ''), ' ', IFNULL(repapematerno, ''), ' ', IFNULL(repnombre, ''))", "(ISNULL(repapepaterno, '')+' '+ISNULL(repapematerno, '')+' '+ISNULL(repnombre, ''))") & " AS representante, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(repdocumento, '') AS documento "
  sSentencia = sSentencia & "FROM tgemp "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "'"
  With porsTemporal
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = sSentencia
    .Open , CONNSTRG & gsNomBDC
    sDireccion = !direccion
    sRepresentante = !representante
    sDocumento = !documento
    .Close
  End With
  
  ' Genero la información
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptceri4ta", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptceri4ta') DROP TABLE #trptceri4ta")
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE #") & "trptceri4ta ("
  sSentencia = sSentencia & "codaux varchar(11) NOT NULL default '', "
  sSentencia = sSentencia & "razaux varchar(80) default NULL, "
  sSentencia = sSentencia & "diraux varchar(80) default NULL, "
  sSentencia = sSentencia & "rucaux varchar(11) default NULL, "
  sSentencia = sSentencia & "rubro varchar(40) default NULL, "
  sSentencia = sSentencia & "moneda varchar(3) default NULL, "
  sSentencia = sSentencia & "impbrur_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impbrun_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impnet_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impir4_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impies_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "import_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impcanr_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impcann_mn decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impbrutopen decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "imprentapen decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "cLetras_1 varchar(80) NOT NULL default '', "
  sSentencia = sSentencia & "cLetras_2 varchar(80) NOT NULL default '')"
  pocnnMain.Execute sSentencia
  
  ' Ingreso los datos al temporal
  ActualizaTemporal
  ' Obtengo la información
  With porstMRp
    If .State = adStateOpen Then .Close
     .Source = "SELECT * "
     .Source = .Source & "FROM " & ps_Prefijo & "trptceri4ta "
     .Source = .Source & "ORDER BY CodAux"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRCerIR4.rpt"
      'Fórmular propias.
      .Formulas(6) = "pEjercicio ='EJERCICIO GRAVABLE " & gsAnoAct & "'"
      .Formulas(7) = "pRucEmpresa = '" & gsRUCEmp & "'"
      .Formulas(8) = "pDireccion=' " & UCase(sDireccion) & "'"
      .Formulas(9) = "Representante='" & UCase(sRepresentante) & "'"
      .Formulas(10) = "RepDocumento='" & sDocumento & "'"
      .Formulas(11) = "pFecha = 'LIMA " & Day(DTPfecha) & " DE " & UCase(Format(DTPfecha, " mmmm ")) & " DEL " & Year(DTPfecha) & "'"
      '.Formulas(8) = "pFecha = 'LIMA " & Day(Date) & " DE " & UCase(Format(Date, " mmmm ")) & " DEL " & Year(Date) & "'"
      ']
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRCerIR4.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption, udFecha, True)
      
      '[Parámetros adicionales.
      .Parameters("pEjercicio") = "EJERCICIO GRAVABLE " & gsAnoAct
      .Parameters("pRucEmpresa") = gsRUCEmp
      .Parameters("pFecha") = "LIMA " & Day(DTPfecha) & " DE " & UCase(Format(DTPfecha, " mmmm ")) & " DEL " & Year(DTPfecha)
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
        'Elimino la tabla temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptceri4ta", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptceri4ta') DROP TABLE #trptceri4ta")

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
   Case 0                              'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
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
   End Select
End Function

'[
Private Sub ActualizaTemporal()
  ' MA
  Dim snLonCad As Integer
  Dim ssCadena As String, sSQL As String
  Dim nRegistros As Long
   
  ' Genero la tabla de saldos de documentos
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocumento", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpdocumento') DROP TABLE #tmpdocumento")
  sSQL = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE ") & ps_Prefijo & "tmpdocumento ("
  sSQL = sSQL & "pdoano char(4) default Null, "
  sSQL = sSQL & "codaux varchar(11) default Null, "
  sSQL = sSQL & "tipdocu char(2) default Null, "
  sSQL = sSQL & "serdoc char(4) default Null, "
  sSQL = sSQL & "nrodoc varchar(10) default Null, "
  sSQL = sSQL & "tpomon char(1) default Null, "
  sSQL = sSQL & "impbru_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impir4_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impnet_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impbru_me decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impir4_me decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impnet_me decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "codcta varchar(16) default Null, "
  sSQL = sSQL & "impcta_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impcta_me decimal(12,2) NOT Null default '0.00')"
  pocnnMain.Execute sSQL
  
  ' Obtengo los importes de provision de documentos
  sSQL = "INSERT INTO " & ps_Prefijo & "tmpdocumento "
  sSQL = sSQL & "SELECT DISTINCT doc.pdoano, doc.codaux, '" & CODTDC_HPR & "' AS tipdocu, doc.serdoc, doc.nrodoc, doc.tpomon, doc.impbru_mn, doc.impir4_mn, doc.impnet_mn, "
  sSQL = sSQL & "doc.impbru_me, doc.impir4_me, doc.impnet_me, cta.codcta , cta.impcta_mn, cta.impcta_me "
  sSQL = sSQL & "FROM cohprdoc doc "
  sSQL = sSQL & "INNER JOIN cohprdoccta cta ON doc.codemp=cta.codemp AND doc.pdoano=cta.pdoano AND doc.codaux=cta.codaux AND doc.serdoc=cta.serdoc AND doc.nrodoc=cta.nrodoc "
  sSQL = sSQL & "WHERE doc.codemp='" & gsCodEmp & "' "
  sSQL = sSQL & "AND doc.pdoano>='" & (Val(gsAnoAct) - 1) & "' "
  sSQL = sSQL & "AND doc.pdoano<='" & gsAnoAct & "' "
  sSQL = sSQL & "AND cta.tpocnc='5' "
  If Trim(txtDato(0).Text) <> "" Then
     sSQL = sSQL & "AND doc.codaux = '" & Trim(txtDato(0).Text) & "' "
  End If
  sSQL = sSQL & "ORDER BY doc.codaux, doc.serdoc, doc.nrodoc"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Genero la tabla de saldos de documentos
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsaldodocu", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#tmpsaldodocu_') DROP TABLE #tmpsaldodocu")
  sSQL = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE ") & ps_Prefijo & "tmpsaldodocu ("
  sSQL = sSQL & "codaux varchar(11) default Null, "
  sSQL = sSQL & "codtdc char(2) default Null, "
  sSQL = sSQL & "serdoc char(4) default Null, "
  sSQL = sSQL & "nrodoc varchar(10) default Null, "
  sSQL = sSQL & "tpomon char(1) default Null, "
  sSQL = sSQL & "impdeb decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "imphab decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "imptcb decimal(7,4) NOT Null default '0.0000') "
  pocnnMain.Execute sSQL
  
  ' Selecciono los documentos pendcientes mn
  sSQL = "INSERT INTO " & ps_Prefijo & "tmpsaldodocu "
  sSQL = sSQL & "SELECT det.codaux, det.codtdc, det.serdoc, det.nrodoc, tmp.tpomon, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END), 0)), 2) AS impdeb, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END), 0)), 2) AS imphab, "
  sSQL = sSQL & "ROUND(AVG(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.imptcb, 1)), 4) AS imptcb "
  sSQL = sSQL & "FROM cocpbdet det "
  sSQL = sSQL & "INNER JOIN " & ps_Prefijo & "tmpdocumento tmp ON det.codcta=tmp.codcta AND det.codaux=tmp.codaux AND det.codtdc=tmp.tipdocu AND det.serdoc=tmp.serdoc AND det.nrodoc=tmp.nrodoc "
  sSQL = sSQL & "WHERE det.codemp='" & gsCodEmp & "' "
  sSQL = sSQL & "AND det.pdoano='" & gsAnoAct & "' "
  sSQL = sSQL & "AND tmp.tpomon='" & TPOMON_NAC & "' "
  sSQL = sSQL & "GROUP BY det.codaux, det.codtdc, det.serdoc, det.nrodoc, tmp.tpomon "
  If ps_Plataforma = pSrvMySql Then
    sSQL = sSQL & "HAVING ROUND(impdeb - imphab, 2) <> 0.00 "
  ElseIf ps_Plataforma = pSrvSql Then
    sSQL = sSQL & "HAVING ROUND(ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END, 0)), 2) - "
    sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END, 0)), 2), 2)<>0.00 "
  End If
  sSQL = sSQL & "ORDER BY det.codaux, det.codtdc, det.serdoc, det.nrodoc"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Selecciono los documentos pendcientes me
  sSQL = "INSERT INTO " & ps_Prefijo & "tmpsaldodocu "
  sSQL = sSQL & "SELECT det.codaux, det.codtdc, det.serdoc, det.nrodoc, tmp.tpomon, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END), 0)), 2) AS impdeb, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END), 0)), 2) AS imphab, "
  sSQL = sSQL & "ROUND(AVG(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.imptcb, 1)), 4) AS imptcb "
  sSQL = sSQL & "FROM cocpbdet det "
  sSQL = sSQL & "INNER JOIN " & ps_Prefijo & "tmpdocumento tmp ON det.codcta=tmp.codcta AND det.codaux=tmp.codaux AND det.codtdc=tmp.tipdocu AND det.serdoc=tmp.serdoc AND det.nrodoc=tmp.nrodoc "
  sSQL = sSQL & "WHERE det.codemp='" & gsCodEmp & "' "
  sSQL = sSQL & "AND det.pdoano='" & gsAnoAct & "' "
  sSQL = sSQL & "AND tmp.tpomon='" & TPOMON_EXT & "' "
  sSQL = sSQL & "GROUP BY det.codaux, det.codtdc, det.serdoc, det.nrodoc, tmp.tpomon "
  If ps_Plataforma = pSrvMySql Then
    sSQL = sSQL & "HAVING ROUND(impdeb - imphab, 2) <> 0.00 "
  ElseIf ps_Plataforma = pSrvSql Then
    sSQL = sSQL & "HAVING ROUND(ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END, 0)), 2) - "
    sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END, 0)), 2), 2)<>0.00 "
  End If
  sSQL = sSQL & "ORDER BY det.codaux, det.codtdc, det.serdoc, det.nrodoc"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Genero la tabla de final de pendientes
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpfinaldocu", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpfinaldocu') DROP TABLE #tmpfinaldocu")
  sSQL = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE #") & "tmpfinaldocu ("
  sSQL = sSQL & "codaux varchar(11) default Null, "
  sSQL = sSQL & "codtdc varchar(2) default Null, "
  sSQL = sSQL & "serdoc varchar(4) default Null, "
  sSQL = sSQL & "nrodoc varchar(10) default Null, "
  sSQL = sSQL & "tpomon char(1) default Null, "
  sSQL = sSQL & "impbru_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impir4_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impnet_mn decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impbru_me decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impir4_me decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "impnet_me decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "imporsal decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "porcepen decimal(6,2) NOT Null default '0.00', "
  sSQL = sSQL & "imporpen decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "rentapen decimal(12,2) NOT Null default '0.00', "
  sSQL = sSQL & "imporcan decimal(12,2) NOT Null default '0.00')"
  pocnnMain.Execute sSQL
  
  ' Selecciono los pendientes finales
  sSQL = "INSERT INTO " & ps_Prefijo & "tmpfinaldocu "
  sSQL = sSQL & "SELECT DISTINCT sal.codaux, sal.codtdc, sal.serdoc, sal.nrodoc, sal.tpomon, "
  sSQL = sSQL & "tmp.impbru_mn, tmp.impir4_mn, tmp.impnet_mn, "
  sSQL = sSQL & "tmp.impbru_me, tmp.impir4_me, tmp.impnet_me, "
  sSQL = sSQL & "ROUND(sal.imphab-sal.impdeb, 2) AS imporsal, "
  sSQL = sSQL & "0.00 AS porcepen, 0.00 AS imporpen, 0.00 AS rentapen, "
  sSQL = sSQL & "0.00 AS imporcan "
  sSQL = sSQL & "FROM " & ps_Prefijo & "tmpsaldodocu sal "
  sSQL = sSQL & "INNER JOIN " & ps_Prefijo & "tmpdocumento tmp ON sal.codaux=tmp.codaux AND sal.codtdc=tmp.tipdocu AND sal.serdoc=tmp.serdoc AND sal.nrodoc=tmp.nrodoc "
  sSQL = sSQL & "ORDER BY sal.codaux, sal.codtdc, sal.serdoc, sal.nrodoc"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Actualizo porcentaje de saldos del documento
  sSQL = "UPDATE " & ps_Prefijo & "tmpfinaldocu "
  sSQL = sSQL & "SET porcepen=ROUND((imporsal*100) / (CASE tpomon WHEN '" & TPOMON_NAC & "' THEN impnet_mn ELSE impnet_me END), 2)"
  pocnnMain.Execute sSQL, nRegistros
  ' Actualizo imprtes pendientes de renta del documento
  sSQL = "UPDATE " & ps_Prefijo & "tmpfinaldocu "
  sSQL = sSQL & "SET imporpen=ROUND((impbru_mn*porcepen) / 100, 2), "
  sSQL = sSQL & "rentapen=ROUND((impir4_mn*porcepen) / 100, 2)"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Elimino y selecciono movimientos cancelados
  sSQL = "DELETE FROM " & ps_Prefijo & "tmpsaldodocu "
  pocnnMain.Execute sSQL, nRegistros
  
  sSQL = "INSERT INTO " & ps_Prefijo & "tmpsaldodocu "
  sSQL = sSQL & "SELECT det.codaux, det.codtdc, det.serdoc, det.nrodoc, tmp.tpomon, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END), 0)), 2) AS impdeb, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END), 0)), 2) AS imphab, "
  sSQL = sSQL & "ROUND(AVG(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.imptcb, 1)), 4) AS imptcb "
  sSQL = sSQL & "FROM cocpbdet det "
  sSQL = sSQL & "INNER JOIN " & ps_Prefijo & "tmpdocumento tmp ON det.codcta=tmp.codcta AND det.codaux=tmp.codaux AND det.codtdc=tmp.tipdocu AND det.serdoc=tmp.serdoc AND det.nrodoc=tmp.nrodoc "
  sSQL = sSQL & "WHERE det.codemp='" & gsCodEmp & "' "
  sSQL = sSQL & "AND det.pdoano='" & gsAnoAct & "' "
  sSQL = sSQL & "AND det.tpopvs='" & TPOPVS_CAN & "' "
  sSQL = sSQL & "GROUP BY det.codaux, det.codtdc, det.serdoc, det.nrodoc "
  sSQL = sSQL & "ORDER BY det.codaux, det.codtdc, det.serdoc, det.nrodoc"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Selecciono los movimientos cancelado
  sSQL = "INSERT INTO " & ps_Prefijo & "tmpfinaldocu "
  sSQL = sSQL & "SELECT DISTINCT sal.codaux, sal.codtdc, sal.serdoc, sal.nrodoc, sal.tpomon, "
  sSQL = sSQL & "tmp.impbru_mn, "
  sSQL = sSQL & "ROUND(CASE WHEN tmp.tpomon ='" & TPOMON_NAC & "' THEN tmp.impir4_mn ELSE (tmp.impir4_me *sal.imptcb) END, 2) AS impir4_mn, "
  sSQL = sSQL & "tmp.impnet_mn, tmp.impbru_me, tmp.impir4_me, tmp.impnet_me, "
  sSQL = sSQL & "0.00 AS imporsal, "
  sSQL = sSQL & "0.00 AS porcepen, 0.00 AS imporpen, 0.00 AS rentapen, "
  '2016-01-29 simulamos datos en campos y si resulta que sale datos en reporte sSQL = sSQL & "0.00 AS porcepen, 1.00 AS imporpen, 2.00 AS rentapen, "
  sSQL = sSQL & "ROUND(CASE WHEN tmp.tpomon ='" & TPOMON_NAC & "' THEN tmp.impbru_mn ELSE (tmp.impbru_me *sal.imptcb) END, 2) AS imporcan "
  sSQL = sSQL & "FROM " & ps_Prefijo & "tmpsaldodocu sal "
  sSQL = sSQL & "INNER JOIN " & ps_Prefijo & "tmpdocumento tmp ON sal.codaux=tmp.codaux AND sal.codtdc=tmp.tipdocu AND sal.serdoc=tmp.serdoc AND sal.nrodoc=tmp.nrodoc "
  sSQL = sSQL & "ORDER BY sal.codaux, sal.codtdc, sal.serdoc, sal.nrodoc"
  pocnnMain.Execute sSQL, nRegistros
  
  ' Documento final
  sSQL = "INSERT INTO " & ps_Prefijo & "trptceri4ta "
  sSQL = sSQL & "SELECT tmp.codaux, aux.razaux, aux.diraux, aux.rucaux, aux.rubro, 'S/.' AS moneda, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE WHEN tmp.impir4_mn> 0 THEN tmp.impbru_mn ELSE 0 END), 0) AS impbrur_mn, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE WHEN tmp.impir4_mn<=0 THEN tmp.impbru_mn ELSE 0 END), 0) AS impbrun_mn, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(tmp.impnet_mn), 0) AS impnet_mn, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(tmp.impir4_mn), 0) AS impir4_mn, "
  sSQL = sSQL & "0.00 AS impies_mn,0.00 AS import_mn, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE WHEN tmp.impir4_mn> 0 THEN tmp.imporcan ELSE 0 END), 0) AS impcanr_mn, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE WHEN tmp.impir4_mn<=0 THEN tmp.imporcan ELSE 0 END), 0) AS impcann_mn, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(tmp.imporpen, 0)), 2) AS impbrutopen, "
  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(tmp.rentapen, 0)), 2) AS imprentapen, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "REPEAT", "REPLICATE") & "(' ', 80) AS cLetras_1, "
  sSQL = sSQL & IIf(ps_Plataforma = pSrvMySql, "REPEAT", "REPLICATE") & "(' ', 80) AS cLetras_2 "
  sSQL = sSQL & "FROM (tmpfinaldocu tmp "
  sSQL = sSQL & "LEFT JOIN tgaux aux ON aux.codemp='" & gsCodEmp & "' AND tmp.CodAux=aux.CodAux) "
  sSQL = sSQL & "GROUP BY tmp.CodAux, aux.RazAux, aux.diraux, aux.RucAux, aux.rubro "
  sSQL = sSQL & "ORDER BY tmp.CodAux"
  pocnnMain.Execute sSQL, nRegistros
  
'' revisar acac parad
  
  ' Genero la tabla de final de pendientes e inserto
'  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsaldopen", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#tmpsaldopen') DROP TABLE #tmpsaldopen")
'  sSQL = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE #") & "tmpsaldopen ("
'  sSQL = sSQL & "codaux varchar(11) default Null, "
'  sSQL = sSQL & "imporpen decimal(12,2) NOT Null default '0.00', "
'  sSQL = sSQL & "rentapen decimal(12,2) NOT Null default '0.00')"
'  pocnnMain.Execute sSQL
'
'  ' Inserto los auxiliares pendientes
'  sSQL = "INSERT INTO " & ps_Prefijo & "tmpsaldopen "
'  sSQL = sSQL & "SELECT codaux, "
'  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(imporpen, 0)), 2) AS imporpen, "
'  sSQL = sSQL & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(rentapen, 0)), 2) AS rentapen "
'  sSQL = sSQL & "FROM " & ps_Prefijo & "tmpfinaldocu "
'  sSQL = sSQL & "GROUP BY codaux "
'  sSQL = sSQL & "ORDER BY codaux"
'  pocnnMain.Execute sSQL, nRegistros
'
'  ' Actualizo los importes en tabla de reportes
'  If ps_Plataforma = pSrvMySql Then
'    sSQL = "UPDATE trptceri4ta rpt, tmpsaldopen pen SET "
'    sSQL = sSQL & "rpt.importepen=pen.imporpen, "
'    sSQL = sSQL & "rpt.rentapen=pen.rentapen "
'    sSQL = sSQL & "WHERE rpt.codaux=pen.codaux"
'  Else
'    sSQL = "UPDATE #trptceri4ta SET "
'    sSQL = sSQL & "importepen=pen.imporpen, "
'    sSQL = sSQL & "rentapen=pen.rentapen "
'    sSQL = sSQL & "FROM #trptceri4ta rpt, #tmpsaldopen pen "
'    sSQL = sSQL & "WHERE rpt.codaux=pen.codaux"
'  End If
'  pocnnMain.Execute sSQL
  
  Set porsTmpRp = New ADODB.Recordset
  With porsTmpRp
      .ActiveConnection = pocnnMain
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Source = "SELECT * "
      .Source = .Source & "FROM " & ps_Prefijo & "trptceri4ta "
      .Source = .Source & "ORDER BY CodAux"
      .Open
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
             pocnnMain.BeginTrans
              .Fields!impbrur_mn = CDec(IIf(IsNull(!impbrur_mn), 0, !impbrur_mn))
              .Fields!impbrun_mn = CDec(IIf(IsNull(!impbrun_mn), 0, !impbrun_mn))
              .Fields!ImpNet_MN = CDec(IIf(IsNull(!ImpNet_MN), 0, !ImpNet_MN))
              .Fields!ImpIR4_MN = CDec(IIf(IsNull(!ImpIR4_MN), 0, !ImpIR4_MN))
              .Fields!ImpIES_MN = CDec(IIf(IsNull(!ImpIES_MN), 0, !ImpIES_MN))
              .Fields!ImpORt_MN = CDec(IIf(IsNull(!ImpORt_MN), 0, !ImpORt_MN))
              .Fields!impcanr_mn = CDec(IIf(IsNull(!impcanr_mn), 0, !impcanr_mn))
              .Fields!impcann_mn = CDec(IIf(IsNull(!impcann_mn), 0, !impcann_mn))
              .Fields!impbrutopen = CDec(IIf(IsNull(!impbrutopen), 0, !impbrutopen))
              ssCadena = gfNumLet(!impbrutopen, TPOMON_NAC)
              ssCadena = "(" & ssCadena & ")"
              .Fields!cletras_1 = ssCadena
              .Fields!imprentapen = CDec(IIf(IsNull(!imprentapen), 0, !imprentapen))
              ssCadena = gfNumLet(!imprentapen, TPOMON_NAC)
              ssCadena = "(" & ssCadena & ")"
              .Fields!cletras_2 = ssCadena
              .Update
             pocnnMain.CommitTrans
              .MoveNext
         Loop
      End If
      .Close
   End With
   Set porsTmpRp = Nothing
   
End Sub

']

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

