VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRTp56CIR 
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
         Format          =   56819713
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
         Value           =   -1  'True
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
         Picture         =   "frmrtp56cir.frx":0000
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
         Picture         =   "frmrtp56cir.frx":01AA
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
         Picture         =   "frmrtp56cir.frx":02F4
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
         Picture         =   "frmrtp56cir.frx":0826
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRTp56CIR"
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

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
'   With porstTGAux
'      .MoveLast
'      txtDato(1).Text = !CodAux
'      .MoveFirst
'      txtDato(0).Text = !CodAux
'   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
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
  Dim sSentencia  As String, sRegistro As String
  Dim porsTemporal As New ADODB.Recordset

  ppHabilitacion False
  
  ' Genero la información
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptceri4ta", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptceri4ta') DROP TABLE #trptceri4ta")
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS ", "CREATE TABLE #") & "trptceri4ta ("
  sSentencia = sSentencia & "CodAux varchar(11) NOT NULL default '', "
  sSentencia = sSentencia & "mespvs char(2) NOT Null default '', "
  sSentencia = sSentencia & "razaux varchar(60) default NULL, "
  sSentencia = sSentencia & "diraux varchar(80) default NULL, "
  sSentencia = sSentencia & "RucAux varchar(11) default NULL, "
  sSentencia = sSentencia & "rubro varchar(40) default NULL, "
  sSentencia = sSentencia & "moneda varchar(3) default NULL, "
  sSentencia = sSentencia & "ImpBru_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpNet_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpIR4_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpIES_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "ImpORt_MN decimal(12,2) NOT NULL default '0.00', "
  sSentencia = sSentencia & "impfac_hpr decimal(7,4) NOT NULL default '0.00')"
  pocnnMain.Execute sSentencia
  ' Seleciono los registros
  sSentencia = "INSERT INTO " & ps_Prefijo & "trptceri4ta "
  sSentencia = sSentencia & "SELECT a.CodAux, a.mespvs, b.razaux, b.diraux, b.RucAux, b.rubro, 'S/.' AS moneda, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpBru_MN), 0), 2) as ImpBru_MN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpNet_MN), 0), 2) as ImpNet_MN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpIR4_MN), 0), 2) as ImpIR4_MN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpIES_MN), 0), 2) as ImpIES_MN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.ImpORt_MN), 0), 2) as ImpORt_MN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AVG(c.impfac_hpr), 0), 4) as impfac_hpr "
  sSentencia = sSentencia & "FROM cohprdoc a "
  sSentencia = sSentencia & "LEFT JOIN tgaux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux "
  sSentencia = sSentencia & "LEFT JOIN cotcbmes c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.mespvs=c.mespvs "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  If Trim(txtDato(0).Text) <> "" Then
     sSentencia = sSentencia & "AND a.CodAux = '" & Trim(txtDato(0).Text) & "' "
  End If
  sSentencia = sSentencia & "GROUP BY a.CodAux, a.mespvs, b.RazAux, b.diraux, b.RucAux, b.rubro "
  sSentencia = sSentencia & "ORDER BY a.CodAux, a.mespvs"
  pocnnMain.Execute sSentencia
  ' Ingreso los datos al temporal
  ActualizaTemporal
  ' Obtengo la información
  With porstMRp
    If .State = adStateOpen Then .Close
     .Source = "SELECT * "
     .Source = .Source & "FROM " & ps_Prefijo & "trptceri4ta "
     .Source = .Source & "ORDER BY codaux, mespvs"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptr56cirhpr.rpt"
      '[ Formulas adicionales del reporte
      sRegistro = FormatNumber(Left(Trim(gsRUCEmp), 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(6) = "mRucEmpresa='" & sRegistro & Mid(Trim(gsRUCEmp), 9) & "'"
      .Formulas(7) = "mDireccion='" & UCase(gsDirEmp) & "'"
      .Formulas(8) = "mDistrito='" & gsLocEmp & "'"
      .Formulas(9) = "mActividad='" & gsGirEmp & "'"
      .Formulas(10) = "mRepresentante='" & UCase(gsRepEmp) & "'"
      sRegistro = FormatNumber(Left(gsRepDNIEmp, 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(11) = "mDniRepresentante='" & sRegistro & Mid(gsRepDNIEmp, 9) & "'"
      .Formulas(12) = "pEjercicio ='EJERCICIO GRAVABLE " & gsAnoAct & "'"
      .Formulas(13) = "pFecha = '" & gsLocEmp & " " & Day(DTPfecha) & " DE " & UCase(Format(DTPfecha, " mmmm ")) & " DEL " & Year(DTPfecha) & "'"
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
  ' Elimino la tabla temporal del reporte
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
            lblDatoDeta(tnIndex).Caption = " " & !RazAux
         End If
      End With
   End Select
End Function

'[
Private Sub ActualizaTemporal()
  ' MA
  Dim sSQL As String
  Dim sAuxiliar As String, sRazon As String, sDireccion As String
  Dim sNroRuc As String, sRubro As String, sMoneda As String
  Dim nContador As Integer
   
  Set porsTmpRp = New ADODB.Recordset
  With porsTmpRp
    .ActiveConnection = pocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = "SELECT * "
    .Source = .Source & "FROM " & ps_Prefijo & "trptceri4ta "
    .Source = .Source & "ORDER BY codaux, mespvs"
    .Open
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        sAuxiliar = !CodAux: sRazon = !RazAux
        sDireccion = IIf(IsNull(!DirAux), "", !DirAux)
        sNroRuc = IIf(IsNull(!RucAux), "", !RucAux)
        sRubro = IIf(IsNull(!rubro), "", !rubro): sMoneda = !moneda
        For nContador = 1 To 12
          If Not .EOF Then
            If !CodAux = sAuxiliar And !mespvs = Format(nContador, "00") Then
              .MoveNext
            Else
              sSQL = "INSERT INTO " & ps_Prefijo & "trptceri4ta (codaux, mespvs, razaux, diraux, RucAux, rubro, moneda) "
              sSQL = sSQL & "VALUES('" & sAuxiliar & "', '" & Format(nContador, "00") & "', "
              sSQL = sSQL & "'" & sRazon & "', "
              sSQL = sSQL & IIf(sDireccion = "", "Null", "'" & sDireccion & "'") & ", "
              sSQL = sSQL & IIf(sNroRuc = "", "Null", "'" & sNroRuc & "'") & ", "
              sSQL = sSQL & IIf(sRubro = "", "Null", "'" & sRubro & "'") & ", "
              sSQL = sSQL & IIf(sMoneda = "", "Null", "'" & sMoneda & "'") & ")"
              pocnnMain.Execute sSQL
            End If
          Else
            sSQL = "INSERT INTO " & ps_Prefijo & "trptceri4ta (codaux, mespvs, razaux, diraux, RucAux, rubro, moneda) "
            sSQL = sSQL & "VALUES('" & sAuxiliar & "', '" & Format(nContador, "00") & "', "
            sSQL = sSQL & "'" & sRazon & "', "
            sSQL = sSQL & IIf(sDireccion = "", "Null", "'" & sDireccion & "'") & ", "
            sSQL = sSQL & IIf(sNroRuc = "", "Null", "'" & sNroRuc & "'") & ", "
            sSQL = sSQL & IIf(sRubro = "", "Null", "'" & sRubro & "'") & ", "
            sSQL = sSQL & IIf(sMoneda = "", "Null", "'" & sMoneda & "'") & ")"
            pocnnMain.Execute sSQL
          End If
        Next nContador
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

