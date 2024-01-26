VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPACM 
   Caption         =   "[Entidad]"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid fgrMeses 
      Bindings        =   "frmPACM.frx":0000
      Height          =   3510
      Left            =   105
      TabIndex        =   0
      Top             =   2640
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6191
      _Version        =   393216
      Rows            =   14
      Cols            =   5
      BackColorSel    =   -2147483640
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   2145
      Left            =   105
      TabIndex        =   1
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3784
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPACM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain_0 As ADODB.Recordset, _
       uorstMain_1 As ADODB.Recordset
Public usConnStrgSele_0 As String, _
       usConnStrgOrde_0 As String, _
       usConnStrgSele_1 As String, _
       usConnStrgWher_1 As String, _
       usConnStrgOrde_1 As String
'       usCOnnStrgWher
Private pnColumnaOrd As Integer

'[Propio del formulario.
'Public uorstCODro As ADODB.Recordset, _
'       uorstTGSvc As ADODB.Recordset
']

'Dim WithEvents MRViewer As MRViewerObject

'Private pocnnMain As ADODB.Connection
'Private porstMRp As ADODB.Recordset

'[Propio del formulario.
Private dnImpIndB As Double
Private porstCOICM As ADODB.Recordset
']

Private Sub dgrMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ppDatosDetalle
End Sub

Private Sub fgrMeses_RowColChange()
'    ppDatosDetalle False
End Sub


Private Sub Form_Load()
 
 '[Recordsets                          'Cambiar.
 
   usConnStrgSele_0 = "SELECT CodCta, DetCta " _
                    & " FROM COCta " _
                    & " WHERE IndMoe=1"   'IndMoe=1 valor para cuenta monetaria
   usConnStrgOrde_0 = " ORDER BY CodCta"
   usConnStrgSele_1 = "SELECT CodCta, " _
                    & " AcuD01_MN, AcuH01_MN, AcuD01_ME, AcuH01_ME, " _
                    & " AcuD02_MN, AcuH02_MN, AcuD02_ME, AcuH02_ME, " _
                    & " AcuD03_MN, AcuH03_MN, AcuD03_ME, AcuH03_ME, " _
                    & " AcuD04_MN, AcuH04_MN, AcuD04_ME, AcuH04_ME, " _
                    & " AcuD05_MN, AcuH05_MN, AcuD05_ME, AcuH05_ME, " _
                    & " AcuD06_MN, AcuH06_MN, AcuD06_ME, AcuH06_ME, " _
                    & " AcuD07_MN, AcuH07_MN, AcuD07_ME, AcuH07_ME, " _
                    & " AcuD08_MN, AcuH08_MN, AcuD08_ME, AcuH08_ME, " _
                    & " AcuD09_MN, AcuH09_MN, AcuD09_ME, AcuH09_ME, " _
                    & " AcuD10_MN, AcuH10_MN, AcuD10_ME, AcuH10_ME, " _
                    & " AcuD11_MN, AcuH11_MN, AcuD11_ME, AcuH11_ME, " _
                    & " AcuD12_MN, AcuH12_MN, AcuD12_ME, AcuH12_ME " _
                    & " FROM COCtaAcu "
   
   Set uocnnMain = New ADODB.Connection
   Set uorstMain_0 = New ADODB.Recordset
   Set uorstMain_1 = New ADODB.Recordset
   Set porstCOICM = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With uorstMain_0
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_0 & usConnStrgOrde_0
'     .CursorLocation = adUseClient   'Es el Default.
        
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open (usConnStrgSele_0 & usConnStrgOrde_0)
'      .Properties("Unique Table").Value = "VTFacCab"
      .Properties("Unique Table").Value = "COCpbCab"
   End With
   With uorstMain_1
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_1
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
'      .Properties("Unique Table").Value = "COCpbDet"
   End With
   With porstCOICM
      .ActiveConnection = uocnnMain
      .Source = "SELECT MesICM, ImpInd " _
              & "FROM CoICM " _
              & "ORDER BY 1"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   porstCOICM.MoveFirst
   porstCOICM.Find "MesICM='" & gfMesAct(gsMesAct) & "'"
   dnImpIndB = porstCOICM!ImpInd
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain_0

   'dgrDetalle.MarqueeStyle = dbgHighlightRow
   'Set dgrDetalle.DataSource = uorstMain_3
End Sub

Private Sub Form_Activate()
   'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   ppDatosGrid
   ppDatosGridMeses
   ppDatosDetalle
'   ppDatosGridDetalle
'   ppDatosDiario
'   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   'gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
On Error GoTo ERR13
'   uorstTGSvc.Close
'   uorstTGArt.Close
'   uorstTGCli.Close
''   uorstVTNroDoc.Close
'   uorstMain_1.Close
   uorstMain_0.Close
   uocnnMain.Close
'   Set uorstTGSvc = Nothing
'   Set uorstTGArt = Nothing
'   Set uorstTGCli = Nothing
'   Set uorstVTNroDoc = Nothing
'   Set uorstMain_1 = Nothing
   Set uorstMain_0 = Nothing
   Set uocnnMain = Nothing
   Exit Sub
ERR13: Resume Next
End Sub

'Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
'
'   On Error GoTo Err
'
'   pnColumnaOrd = ColIndex
''   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
'   txtBuscar = ""
'
'   usConnStrgOrde_0 = "ORDER BY "
''   Select Case pnColumnaOrd            'Cambiar.
''   Case 1, 2, 3
''      usConnStrgOrde_0 = usConnStrgOrde_0 & "1, 2, 3"
''   Case Else
'      usConnStrgOrde_0 = usConnStrgOrde_0 & pnColumnaOrd + 1
''   End Select
'   With uorstMain_0
'      .Close
'      .Source = usConnStrgSele_0 & usConnStrgOrde_0
'      .Open
'   End With
'   Set dgrMain.DataSource = uorstMain_0
'   ppDatosGrid
'
'   Exit Sub
'
'Err: gpErrores
'
'End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   
    Select Case KeyCode
        Case vbKeyHome
            uorstMain_0.MoveFirst
        Case vbKeyEnd
            uorstMain_0.MoveLast
    End Select
    
End Sub

'Private Sub txtBuscar_Change()
'   On Error GoTo Err
'
'   Dim dsCriterio As String
'   Dim dvRegistroActual As Variant
'
'   With uorstMain_0
'      dvRegistroActual = .Bookmark
'
''[ARREGLAR: Búsqueda con distintos tipos de columna.
'      Select Case VarType(.Fields(pnColumnaOrd))
'      Case vbString
'         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
'      Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
'         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
''     Case vbDate
''         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'      End Select
'      .Find dsCriterio, , , 1
'      If .EOF = True Then
'         .Bookmark = dvRegistroActual
'      End If
'   End With
'']ARREGLAR.
'
'   Exit Sub
'Err:
'   If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
'      uorstMain_0.Bookmark = dvRegistroActual
'   Else
'      gpErrores
'   End If
'End Sub
'
Public Sub ppDatosGrid()               'Cambiar Datos Grid.
   
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = "Cod.Cta."
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("CodCta").DefinedSize + 2)
        Case 1
            .Item(dnNum).Caption = "Descripción"
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("DetCta").DefinedSize + 2)
         Case 2
            .Item(dnNum).Caption = "         Tipo de Moneda"
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("cTpoMon").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgCenter
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
   
End Sub


Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
'   cmdNuevo.Enabled = taOpciones(0)
'   cmdEliminar.Enabled = taOpciones(1)
'   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property
Public Sub ppDatosGridMeses()
    Dim dnNum    As Integer
         
    With fgrMeses
        .Clear
        For dnNum = 0 To .Cols - 1
            Select Case dnNum
                Case 0: .ColWidth(dnNum) = 1500
                        .TextMatrix(0, dnNum) = "     Meses\Saldos"
                Case 1: .ColWidth(dnNum) = 1200
                        .ColAlignment(dnNum) = 7
                        .TextMatrix(0, dnNum) = "Importe"
                Case 2: .ColWidth(dnNum) = 1200
                        .ColAlignment(dnNum) = 7
                        .TextMatrix(0, dnNum) = "Fact./Aj."
                Case 3: .ColWidth(dnNum) = 1200
                        .ColAlignment(dnNum) = 7
                        .TextMatrix(0, dnNum) = "Importe Aj."
                Case 4: .ColWidth(dnNum) = 1200
                        .ColAlignment(dnNum) = 7
                        .TextMatrix(0, dnNum) = "Saldo"
            End Select
        Next dnNum
        
        For dnNum = 0 To .Rows - 1
            Select Case dnNum
                Case 1:  .TextMatrix(dnNum, 0) = "Enero"
                Case 2:  .TextMatrix(dnNum, 0) = "Febrero"
                Case 3:  .TextMatrix(dnNum, 0) = "Marzo"
                Case 4:  .TextMatrix(dnNum, 0) = "Abril"
                Case 5:  .TextMatrix(dnNum, 0) = "Mayo"
                Case 6:  .TextMatrix(dnNum, 0) = "Junio"
                Case 7:  .TextMatrix(dnNum, 0) = "Julio"
                Case 8:  .TextMatrix(dnNum, 0) = "Agosto"
                Case 9:  .TextMatrix(dnNum, 0) = "Setiembre"
                Case 10: .TextMatrix(dnNum, 0) = "Octubre"
                Case 11: .TextMatrix(dnNum, 0) = "Noviembre"
                Case 12: .TextMatrix(dnNum, 0) = "Diciembre"
                Case 13: .TextMatrix(dnNum, 0) = "TOTALES"
            End Select
        Next dnNum
        .Row = Val(gfMesAct(gsMesAct)) + 1
        .SetFocus
    End With

End Sub

Public Sub ppDatosDetalle()

   Dim dnfil As Integer
   Dim dnSum As Double, dnSumAj As Double
   Dim cTotalMN As Double
   Dim dnImpDif As Double, dnFactAj As Double
       
   dnSum = 0#: dnSumAj = 0#: cTotalMN = 0#
   With uorstMain_1
       If .RecordCount > 0 And uorstMain_0.RecordCount > 0 Then
          .MoveFirst
          .Find "CodCta='" & uorstMain_0!CodCta & "'"
         If Not .EOF Then
             For dnfil = 1 To 12
                 If dnfil <= CInt(gsMesAct) Then
                     With fgrMeses
                         porstCOICM.MoveFirst
                         porstCOICM.Find "MesICM='" & Format(dnfil, "00") & "'"
                         dnFactAj = Round(dnImpIndB / IIf(porstCOICM!ImpInd = 0, 1, porstCOICM!ImpInd), 3)
                         dnImpDif = uorstMain_1.Fields("AcuD" & Format(dnfil, "00") & "_MN") - uorstMain_1.Fields("AcuH" & Format(dnfil, "00") & "_MN")
                         .TextMatrix(dnfil, 1) = Format(dnImpDif, FORMATO_NUM_1)
                         dnSum = dnSum + Format(.TextMatrix(dnfil, 1), FORMATO_NUM_2)
                         .TextMatrix(dnfil, 2) = Format(dnFactAj, FORMATO_NUM_2)
                         .TextMatrix(dnfil, 3) = Format(dnImpDif * dnFactAj, FORMATO_NUM_1)
                         dnSumAj = dnSumAj + Format(.TextMatrix(dnfil, 3), FORMATO_NUM_2)
                         .TextMatrix(dnfil, 4) = Format(.TextMatrix(dnfil, 1) - .TextMatrix(dnfil, 3), FORMATO_NUM_1)
                         cTotalMN = cTotalMN + Format(.TextMatrix(dnfil, 4), FORMATO_NUM_2)
                     End With
                  Else
                     With fgrMeses
                         .TextMatrix(dnfil, 1) = "0.00"
                         .TextMatrix(dnfil, 2) = "0.00"
                         .TextMatrix(dnfil, 3) = "0.00"
                         .TextMatrix(dnfil, 4) = "0.00"
                     End With
                  End If
             Next dnfil
             With fgrMeses
                 .TextMatrix(13, 1) = Format(dnSum, FORMATO_NUM_1)
                 .TextMatrix(13, 3) = Format(dnSumAj, FORMATO_NUM_1)
                 .TextMatrix(13, 4) = Format(cTotalMN, FORMATO_NUM_1)
             End With
         Else
            For dnfil = 1 To fgrMeses.Rows - 3
            With fgrMeses
                .TextMatrix(dnfil, 1) = "0.00"
                .TextMatrix(dnfil, 2) = "0.00"
                .TextMatrix(dnfil, 3) = "0.00"
                .TextMatrix(dnfil, 4) = "0.00"
            End With
           Next dnfil
           With fgrMeses
               .TextMatrix(13, 1) = "0.00"
               .TextMatrix(13, 3) = "0.00"
               .TextMatrix(13, 4) = "0.00"
           End With
         End If
       Else
            For dnfil = 1 To fgrMeses.Rows - 3
            With fgrMeses
                .TextMatrix(dnfil, 1) = "0.00"
                .TextMatrix(dnfil, 2) = "0.00"
                .TextMatrix(dnfil, 3) = "0.00"
                .TextMatrix(dnfil, 4) = "0.00"
            End With
           Next dnfil
           With fgrMeses
               .TextMatrix(13, 1) = "0.00"
               .TextMatrix(13, 3) = "0.00"
               .TextMatrix(13, 4) = "0.00"
           End With
       End If
   End With
End Sub
