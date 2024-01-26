VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPAsFmtTran 
   Caption         =   "Transferencia de Tipo Asiento - Diario"
   ClientHeight    =   6270
   ClientLeft      =   8460
   ClientTop       =   1260
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   6705
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   450
      Left            =   2010
      TabIndex        =   20
      Top             =   5610
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   450
      Left            =   3690
      TabIndex        =   19
      Top             =   5610
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.TextBox txtDato 
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Text            =   "9201"
      Top             =   390
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   4215
      Picture         =   "frmPAsFmtTran.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   2
      Left            =   4185
      Picture         =   "frmPAsFmtTran.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3105
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDato 
      Height          =   300
      Index           =   2
      Left            =   120
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3105
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   4185
      Picture         =   "frmPAsFmtTran.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDato 
      Height          =   300
      Index           =   1
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   465
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1110
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1830
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   3765
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ingrese Diario"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   2805
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
      Left            =   570
      TabIndex        =   16
      Top             =   390
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ajuste por Documento"
      ForeColor       =   &H80000002&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ajuste por Cuenta"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ajuste por Cuenta + Auxiliar"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   5
      Left            =   135
      TabIndex        =   13
      Top             =   3525
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblTexto 
      Caption         =   "Flujo de Caja Perdida :"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   2850
      Visible         =   0   'False
      Width           =   2805
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
      Left            =   540
      TabIndex        =   11
      Top             =   3105
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Label lblTexto 
      Caption         =   "Flujo de Caja Ganancia :"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   2265
      Visible         =   0   'False
      Width           =   2805
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
      Left            =   540
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   3660
   End
End
Attribute VB_Name = "frmPAsFmtTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
'Public porstCoTCbMes As ADODB.Recordset
Public porstCodro As ADODB.Recordset
'Public porstCoFjo As ADODB.Recordset
Public porstCOmCpbCab_ok As ADODB.Recordset
Public rcTPOGNR_DCA As String

Private Sub Form_Load()
   pgbProceso(0).Value = 0
'''   pgbProceso(1).Value = 0
  
  'Abrir Tablas.
   
   Set pocnnMain = New ADODB.Connection
'''   Set porstCoTCbMes = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
'''   Set porstCoFjo = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With

   
   '[ Cargo los mensajes de botones
   Dim nElemento As Integer
   ReDim aLabel(6, 2)
   For nElemento = 0 To UBound(aLabel, 1) - 1
     aLabel(nElemento, 0) = Choose(nElemento + 1, "Ingrese Diario", "Ajuste por Documento", "Ajuste por Cuenta", "Flujo de Caja Ganancia", "Flujo de Caja Perdida", "Ajuste por Cuenta + Auxiliar")
     aLabel(nElemento, 1) = Choose(nElemento + 1, "Enter Journal", "Adjustment for Document", "Adjustment for Account", "Cash Flow Profit", "Cash Flow Loss", "Adjustment for Account + Auxiliary")
   Next nElemento
   cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
   CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']


   With porstCodro
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
      .Source = .Source & "FROM CODro "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   
   '2008-jun-19
   'elije exportacion
   ' Abro el recordset de formatos de comprobantes
   Set porstCOmCpbCab_ok = New ADODB.Recordset
   With porstCOmCpbCab_ok
       'If .State = adStateOpen Then .Close
       .ActiveConnection = pocnnMain
       'Genero la sentencia de seleccion cabecera de comprobantes
       .Source = "SELECT Copia, CodDro, NroCpb, FehCpb, " & Choose(gsIdioma, "glocpb", "glocpbx") & " "
'       .Source = "SELECT Copia, codemp, pdoano, MesPvs, CodDro, NroCpb, FehCpb, glocpb, glocpbx "
'       .Source = .Source & "Copia "
'       .Source = .Source & "TpoGnr, IndNCu, IndAnu, "
'       .Source = .Source & "UsrCre, FyHCre "
       .Source = .Source & "FROM COmaCpbCab "
       .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
       .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'         .Source = .Source & "AND CodDro='9999'"
       .CursorType = adOpenDynamic
       .LockType = adLockOptimistic
       .Open
    End With
    ' Propiedades del Flex
    With MSFlexGrid1
        .Rows = 2
        .cols = 1
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionFree
    End With
     ' rellena la grilla con los datos del recordset. _
      Le pasa el nímero o index del campo queactuará como CheckBox
    Call Cargar_FlexGrid(MSFlexGrid1, 0, porstCOmCpbCab_ok)

End Sub
'===========================================================
'Sub que carga los registros en la Grilla
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Cargar_FlexGrid(FlexGrid As Object, _
                            NumeroCampo As Integer, _
                            objRst As ADODB.Recordset)
                               
On Local Error GoTo ErrSub
                               
Dim c As Integer
Dim fila As Integer
Dim AnchoCol() As Single
Dim TempAnchoCol As Single
  
    With FlexGrid
        ' deshabilita el repintado para que sea mas rápido
        .Redraw = False
        'Cantidad de filas y columnas
        .Rows = 1
        .cols = objRst.Fields.Count
    End With
       
    'Redimensiona el Array a la cantidad de campos del recordset
    ReDim AnchoCol(0 To objRst.Fields.Count - 1)
       
    'Recorre las columnas
    For c = 0 To objRst.Fields.Count - 1
        'Añade el título del campo al encabezado de columna
        FlexGrid.TextMatrix(0, c) = objRst.Fields(c).Name
        'Guarda el ancho del campo en la matriz
        AnchoCol(c) = TextWidth(objRst.Fields(c).Name)
        Select Case c
        Case 0
            FlexGrid.TextMatrix(0, c) = "Copia"
            AnchoCol(c) = 200
        Case 1
            FlexGrid.TextMatrix(0, c) = "Diario"
        Case 2
            FlexGrid.TextMatrix(0, c) = "Comprob."
        Case 3
            FlexGrid.TextMatrix(0, c) = "Fecha"
        Case 4
            FlexGrid.TextMatrix(0, c) = "Glosa"
       End Select

    Next c
    fila = 1
       
    'Recorre todos los registros del recordset
    Do While Not objRst.EOF
        ' Añade una nueva fila
        FlexGrid.Rows = FlexGrid.Rows + 1
        For c = 0 To objRst.Fields.Count - 1
               
            'Si el valor no es nulo
            If Not IsNull(objRst.Fields(c).Value) Then
               ' si la columna es el campo de tipo CheckBox ...
               If c = NumeroCampo Then
                    With FlexGrid
                        .row = fila ' se posiciona en la fila
                       .Col = c '  .. en la columna
                       ' cambia la fuente para esta celda
                        .CellFontName = "Wingdings"
                        .CellFontSize = 14
                        .CellAlignment = flexAlignCenterCenter
                        ' edita la celda
                        If objRst(NumeroCampo).Value = True Then
                            .TextMatrix(fila, NumeroCampo) = Chr(254) ' false
                        Else
                            .TextMatrix(fila, NumeroCampo) = Chr(168) ' true
                        End If
                    End With
                       
               Else
                   'Agrega el registro en la fila y columna específica
                   FlexGrid.TextMatrix(fila, c) = objRst.Fields(c).Value
                    ' Almacena el ancho
                   TempAnchoCol = TextWidth(objRst.Fields(c).Value)
               End If
                              
               If AnchoCol(c) < TempAnchoCol Then
                  AnchoCol(c) = TempAnchoCol ' nuevo ancho
               End If
            End If
        Next
        ' Siguiente registro
        objRst.MoveNext
        fila = fila + 1 'Incrementa la posición de la fila actual
    Loop
  
    ' Establece los ancho máximos de columna
    For c = 0 To FlexGrid.cols - 1
        FlexGrid.ColWidth(c) = AnchoCol(c) + 240
    Next
    FlexGrid.ColWidth(0) = 800
    ' vuelve a habilitar el redraw
    FlexGrid.Redraw = True
Exit Sub
  
'Error
ErrSub:
MsgBox Err.Description, vbCritical
FlexGrid.Redraw = True
End Sub

'===========================================================

Private Sub cmdAceptar_Click()
'?teo cambiar numero de tranferencia
 'rcTPOGNR_DCA = "9"
'?teo confirm numero 2008/jul/21
rcTPOGNR_DCA = TPOGNR_DRO

  Dim dnContador As Integer
  On Error GoTo Err
   
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
 '[Propio del formulario.
  For dnContador = 0 To txtDato.Count - 1
    If txtDato.Item(dnContador).Text = "" Then
      If dnContador = 0 Then
      MsgBox TEXT_6002, vbCritical
      txtDato(dnContador).SetFocus
      Exit Sub
      End If
    End If
  Next dnContador
   
   If gnIndMNE <> INDMNE_ACT Then
      MsgBox Choose(gsIdioma, "La Empresa trabaja sólo con una Moneda", "The Company works only one Currency"), vbExclamation
      Exit Sub
   End If
   pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
  ' pgbProceso(1).Value = 0: pgbProceso(1).Min = 0
  ' pgbProceso(2).Value = 0: pgbProceso(2).Min = 0
   
   pocnnMain.BeginTrans                'INICIA TRANSACCION.
  
  'Paso 1 : Elimino los comprobantes de ajuste del mes
  
   '2008/jun/21 tendra que borrarse de uno en uno
   'pocnnMain.Execute "DELETE FROM cocpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(rcTPOGNR_DCA) & " AND MesPvs='" & gsMesAct & "'"
   
   
  'Paso 2 : Generacion de Ajustes por Documento
'   ppAjuste_Documento
    ppAsFmt_Tranfe
''  'Paso 3 : Generacion de Ajustes por Cuenta
''   ppAjuste_SaldoCuenta
''  'Paso 4 : Generacion de Ajustes por Cuenta
''   ppAjuste_Auxiliar

   pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
   
   MsgBox TEXT_8008, vbInformation
  
   Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description



End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
'''   cmdAceptar.Enabled = False
'''   cmdSalir.Enabled = True
'''   cmdSalir.SetFocus
End Sub
Private Sub ppAsFmt_Tranfe()

'variable de contro loop
Dim xChanca As Integer
xChanca = 0
'xChanca = 1 no chanca informacion salta procso

'variable para busqueda de historico
Dim mFind As String
'2008/jun/22
Dim uorstCOmCpbCab_ctr As ADODB.Recordset
Set uorstCOmCpbCab_ctr = New ADODB.Recordset
With uorstCOmCpbCab_ctr
   .ActiveConnection = pocnnMain
   'Genero la sentencia de seleccion cabecera de comprobantes
   'es MesPvs y no MesPvs_f, por que debe verificar
   'que existe el formato en el mes indicado
   .Source = "SELECT codemp_f, pdoano_f, MesPvs_f, CodDro_f, NroCpb_f, codemp, pdoano, MesPvs, CodDro, NroCpb,"
   .Source = .Source & "CONCAT(codemp_f, pdoano_f, MesPvs, CodDro_f, NroCpb_f) as xFind,"
   .Source = .Source & "UsrCre, FyHCre "
   .Source = .Source & "FROM COmaCpbCab_ctr "
   .Source = .Source & "WHERE codemp_f='" & gsCodEmp & "' "
   .Source = .Source & "AND pdoano_f='" & gsAnoAct & "' "
  '   .Source = .Source & "Order by codemp_f + pdoano_f + MesPvs_f + CodDro_f + NroCpb_f"
  '   .Source = .Source & "AND Copia=-1"
   .CursorType = adOpenDynamic
   .LockType = adLockReadOnly
   .Open
End With

Dim porstCOmCpbCab_ctr As ADODB.Recordset
Set porstCOmCpbCab_ctr = New ADODB.Recordset
With porstCOmCpbCab_ctr
   .ActiveConnection = pocnnMain
   'Genero la sentencia de seleccion cabecera de comprobantes
   .Source = "SELECT codemp_f, pdoano_f, MesPvs_f, CodDro_f, NroCpb_f, codemp, pdoano, MesPvs, CodDro, NroCpb, "
   '.Source = .Source & "codemp_f + pdoano_f + MesPvs_f + CodDro_f + NroCpb_f as xFind,"
   .Source = .Source & "UsrCre, FyHCre "
   .Source = .Source & "FROM COmaCpbCab_ctr "
'   .Source = .Source & "WHERE codemp_f='" & gsCodEmp & "' "
'   .Source = .Source & "AND pdoano_f='" & gsAnoAct & "' "
 .Source = .Source & "WHERE codemp_f='" & gsCodEmp & "' "
 .Source = .Source & "AND pdoano_f='" & gsAnoAct & "' "
   .Source = .Source & "AND CodDro_f=''"
   '.Source = .Source & "Order by codemp_f + pdoano_f + MesPvs_f + CodDro_f + NroCpb_f "
'   .Source = .Source & "AND Copia=-1"
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .Open
End With

'fin de cambios
'2008/jun/21 insertando tipo de cambio
Dim uorstTGTCb As ADODB.Recordset
Set uorstTGTCb = New ADODB.Recordset
With uorstTGTCb
  .ActiveConnection = pocnnMain
  .Source = "SELECT a.FehTCb, a.ImpTCb_Cpr, a.ImpTCb_Vta "
  .Source = .Source & "FROM TGTCb a "
  .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
'     .CursorLocation = adUseClient   'Es el Default.
   .CursorType = adOpenDynamic
   .LockType = adLockReadOnly
   .Open
End With
'fin de insercion
'tipo de cambio

'2008/jun/21 insertando plan de cuentas
 Dim uorstCoCta As ADODB.Recordset
 Set uorstCoCta = New ADODB.Recordset
With uorstCoCta
 .ActiveConnection = pocnnMain
 .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta, "
 .Source = .Source & "TpoTCb, TpoAnl, IndAjd, IndCCo, IndDoc, IndFjo, CodCCo_Def, "
 .Source = .Source & "CodCta_AjD_Deb, CodCta_AjD_Hab, CodCCo_AjD_Deb, CodCCo_AjD_Hab "
 .Source = .Source & "FROM COCta "
 .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
 .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
 .Source = .Source & "AND TpoCta=" & TPOCTA_TRA & " "
 .Source = .Source & "AND EstCta='" & ESTCTA_ACT & "'"
 '     .CursorLocation = adUseClient   'Es el Default.
 .CursorType = adOpenDynamic
 .LockType = adLockReadOnly
 .Open
End With
'fin de insercion

Dim xTpoTcb As Double
Dim xImpTCb As Double
Dim xTpomon As Double
Dim nImpMN As Double, nImpME As Double
Dim xdcValIndDoc As Double
Dim xFecha As String
Dim xDia As String
Dim sCenCosto As String, sCodCCo_Ajd As String
Dim sTpoCtb_Ajd As String, sTpoMon_Ajd As String
Dim nImpTCb_Ajd As Double, nImporte As Double
Dim nImpMN_Ajd As Double, nImpME_Ajd As Double

Dim sNroComprobante As String
Dim nNroItem As Integer, nContador As Integer

Static porstCOCpbCab As ADODB.Recordset
Static porstCOCpbDet As ADODB.Recordset
Static porstCoCpbAjD As ADODB.Recordset
Static porstUltCoCpb  As ADODB.Recordset

Static porstCOmCpbCab As ADODB.Recordset
Static porstCOmCpbDet As ADODB.Recordset

Set porstCOCpbCab = New ADODB.Recordset
Set porstCOCpbDet = New ADODB.Recordset
Set porstCoCpbAjD = New ADODB.Recordset
Set porstUltCoCpb = New ADODB.Recordset

Set porstCOmCpbCab = New ADODB.Recordset
Set porstCOmCpbDet = New ADODB.Recordset
' Abro el recordset de formatos de comprobantes
With porstCOmCpbCab
   'If .State = adStateOpen Then .Close
   .ActiveConnection = pocnnMain
   'Genero la sentencia de seleccion cabecera de comprobantes
   .Source = "SELECT codemp, pdoano, MesPvs, CodDro, NroCpb, FehCpb, glocpb, glocpbx, "
   .Source = .Source & "TpoGnr, IndNCu, IndAnu, "
   .Source = .Source & "UsrCre, FyHCre "
   .Source = .Source & "FROM COmaCpbCab "
   .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
   .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
   .Source = .Source & "AND Copia=-1"
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .Open
End With
' Abro el recordset de grabacion de la cabecera de comprobante
With porstCOCpbCab
   .ActiveConnection = pocnnMain
   'Genero la sentencia de seleccion cabecera de comprobantes
   .Source = "SELECT codemp, pdoano, CodDro, NroCpb, FehCpb, GloCpb, GloCpbx, MesPvs, "
   .Source = .Source & "TpoGnr, IndNCu, IndAnu, "
   .Source = .Source & "UsrCre, FyHCre "
   .Source = .Source & "FROM COCpbCab "
   .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
   .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
   .Source = .Source & "AND CodDro=''"
   .CursorType = adOpenDynamic
   .LockType = adLockOptimistic
   .Open
End With
Do While Not porstCOmCpbCab.EOF
'While Not porstCOmCpbCab.EOF
  ' Obtengo el numero e inserto la cabecera del comprobante
  With porstUltCoCpb
     If .State = adStateOpen Then .Close
     .ActiveConnection = pocnnMain
     .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
     .Source = .Source & "FROM COCpbCab "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
     .Source = .Source & "AND CodDro='" & porstCOmCpbCab!coddro & "'"
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
     sNroComprobante = !cUltNroCpb
     .Close
  End With
  sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
      '2008/jun/22 control de duplicado
    'de copia formato
    With uorstCOmCpbCab_ctr
       If .EOF Then
        With porstCOmCpbCab_ctr
            .AddNew
            !codemp_f = gsCodEmp
            !pdoano_f = gsAnoAct
            !mespvs_f = porstCOmCpbCab!mespvs
            !CodDro_f = porstCOmCpbCab!coddro
            !NroCpb_f = porstCOmCpbCab!NroCpb
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !mespvs = gsMesAct
            !coddro = porstCOmCpbCab!coddro
            !NroCpb = sNroComprobante
            !UsrCre = gsAbvUsr
            !FyHCre = Now
            .Update
        End With
       Else
           xChanca = 0
          .MoveFirst
           mFind = "xFind='" & _
                  porstCOmCpbCab!codemp & porstCOmCpbCab!pdoano & gsMesAct & _
                  porstCOmCpbCab!coddro & porstCOmCpbCab!NroCpb & _
                  "'"
                  'Cambio porstCOmCpbCab!mespvs por gsMesAct
                  'pues debe revisar si existe en el mes el
                  'comprobante
                  
          .Find mFind
          'si no existe es nueva transferencia
          If .EOF Then
           ' MsgBox TEXT_8006, vbExclamation
            With porstCOmCpbCab_ctr
                .AddNew
                !codemp_f = gsCodEmp
                !pdoano_f = gsAnoAct
                !mespvs_f = porstCOmCpbCab!mespvs
                !CodDro_f = porstCOmCpbCab!coddro
                !NroCpb_f = porstCOmCpbCab!NroCpb
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !mespvs = gsMesAct
                !coddro = porstCOmCpbCab!coddro
                !NroCpb = sNroComprobante
                !UsrCre = gsAbvUsr
                !FyHCre = Now
                .Update
            End With
          Else
             'MsgBox "Existe Comprobante ", vbExclamation
    '         xdcValIndDoc = uorstCOCta!IndDoc
    
            If MsgBox("Comprobante: " & " " & Trim(!coddro) & " (" & Trim(!NroCpb) & "), ya esta transferido, desea volverlo a pasar ?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
            'borro comprobante
            'cambio nro comprobante + 1, al numero de mi historico
             pocnnMain.Execute "DELETE FROM cocpbCab WHERE codemp='" & gsCodEmp & _
             "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(rcTPOGNR_DCA) & _
             " AND MesPvs='" & gsMesAct & "'" & _
             " AND CodDro='" & uorstCOmCpbCab_ctr!coddro & "'" & _
             " AND NroCpb='" & uorstCOmCpbCab_ctr!NroCpb & "'"
               'para que agarre el comprobante nro
               'comprante a chancar
              sNroComprobante = uorstCOmCpbCab_ctr!NroCpb

            Else
                'no chanca informacion salta procso
               xChanca = 1

'              pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
'              Exit Sub
            End If

    
          End If
       End If
          '2014-03-21 si esta vacio no puede hacer esto.MoveFirst
          '.MoveFirst
          If Not .EOF Then
          .MoveFirst
          End If
    End With
    'fin proceso
If xChanca = 0 Then
  With porstCOCpbCab
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !mespvs = gsMesAct
    'a que diario ira
    'los formatos deben reflejar el diario
    'que mantendran en la copia
    'debe ser el mismo diario
    !coddro = porstCOmCpbCab!coddro 'txtDato(0).Text
    !NroCpb = sNroComprobante
    xDia = gfCeros(Day(porstCOmCpbCab!FehCpb), 2, 0, "0")
    xFecha = gfUltDia(xDia & "/" & gsMesAct & "/" & gsAnoAct)
    If Val(xDia) <= Val(Left(xFecha, 2)) Then
         xFecha = xDia & "/" & gsMesAct & "/" & gsAnoAct
    End If
    '2008/jun/21 busca fecha tipo cambio
    With uorstTGTCb
     If .RecordCount <> 0 Then
       .MoveFirst
    '            .Find "FehTCb = '" & IIf(dcValIndDoc = INDDOC_ACT, IIf(optTpoPvs(1).Value, frmMCpbDet.dtpFehOpe, frmMCpbDet.dtpFeEDoc), frmMCpbDet.dtpFehOpe) & "'"
    '               .Find "FehTCb = '" & xDia & "/" & gsMesAct & "/" & gsAnoAct & "'"
       .Find "FehTCb = '" & xFecha & "'"
       If .EOF Then
         MsgBox TEXT_9015 & " Comprobante : " & porstCOmCpbCab!coddro & "-" & porstCOmCpbCab!NroCpb, vbExclamation
       Else
           xImpTCb = Format(IIf(xTpoTcb = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
       End If
       ']
     Else
    '            frmMCpbDet.txtDato(8).Text = Format(0, FORMATO_NUM_2)
     End If
    End With
    'fin de fecha
'    !FehCpb = xDia & "/" & gsMesAct & "/" & gsAnoAct
'2008/jul/21 rcs error en creacion fecha
    !FehCpb = xFecha
    !glocpb = porstCOmCpbCab!glocpb
    !glocpbx = porstCOmCpbCab!glocpbx
    !tpognr = rcTPOGNR_DCA
    !IndNCu = INDNCU_FAL
    !IndAnu = INDANU_FAL
    !UsrCre = gsAbvUsr
    !FyHCre = Now
    .Update
    End With
    nNroItem = 0
    ' Abro el recordset de grabacion del detalle del formato
    With porstCOmCpbDet
       If .State = adStateOpen Then .Close
       .ActiveConnection = pocnnMain
       'Genero la sentencia de seleccion detalles de comprobantes
       .Source = "SELECT codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
       .Source = .Source & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, gloitex, TpoCtb, TpoPvs, IndFjo_Det, "
       .Source = .Source & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, "
       .Source = .Source & "UsrCre, FyHCre "
       .Source = .Source & "FROM COmaCpbDet "
       .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
       .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
       .Source = .Source & "AND CodDro='" & porstCOmCpbCab!coddro & "' "
       .Source = .Source & "AND NroCpb='" & porstCOmCpbCab!NroCpb & "' "
       .Source = .Source & "ORDER BY MesPvs, NroIte"
       .CursorType = adOpenDynamic
       .LockType = adLockOptimistic
       .Open
    End With
    ' Abro el recordset de grabacion del detalle del comprobante
    With porstCOCpbDet
       If .State = adStateOpen Then .Close
       .ActiveConnection = pocnnMain
       'Genero la sentencia de seleccion detalles de comprobantes
       .Source = "SELECT codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
       .Source = .Source & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, gloitex, TpoCtb, TpoPvs, IndFjo_Det, "
       .Source = .Source & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, "
       .Source = .Source & "UsrCre, FyHCre "
       .Source = .Source & "FROM COCpbDet "
       .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
       .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
       .Source = .Source & "AND CodDro=''"
       .CursorType = adOpenDynamic
       .LockType = adLockOptimistic
       .Open
    End With
    Do While Not porstCOmCpbDet.EOF
        nNroItem = nNroItem + 1
        '2008/jun/21 busca cuenta
        With uorstCoCta
           .MoveFirst
           .Find "CodCta='" & porstCOmCpbDet!codcta & "'"
           If .EOF Then
              MsgBox TEXT_8006, vbExclamation
           Else
              xdcValIndDoc = uorstCoCta!IndDoc
           End If
        End With
        'fin de busqueda
        '2008/jun/06 conversion a otra moneda
        If porstCOmCpbDet!TpoTcb = "V" Then
           xTpoTcb = 0
        Else
           xTpoTcb = 1
        End If
        With uorstTGTCb
         If .RecordCount <> 0 Then
             .MoveFirst
             'TPOPVS_CAN, =optTpoPvs(1).Value
             '            .Find "FehTCb = '" & IIf(xdcValIndDoc = INDDOC_ACT, IIf(porstCOmCpbDet!TpoPvs = TPOPVS_CAN, " xDia & " / " & gsMesAct & " / " & gsAnoAct & ", frmMCpbDet.dtpFeEDoc), frmMCpbDet.dtpFehOpe) & "'"
             '            .Find "FehTCb = '" & IIf(xdcValIndDoc = INDDOC_ACT, IIf(porstCOmCpbDet!TpoPvs = TPOPVS_CAN, " xDia & " / " & gsMesAct & " / " & gsAnoAct & ", frmMCpbDet.dtpFeEDoc), frmMCpbDet.dtpFehOpe) & "'"
                '2008/jul/21 rcs error en creacion fecha
                '            .Find "FehTCb = '" & xDia & "/" & gsMesAct & "/" & gsAnoAct & "'"
              .Find "FehTCb = '" & xFecha & "'"
            
             If .EOF Then
               MsgBox TEXT_9015 & " Comprobante : " & porstCOmCpbCab!coddro & "-" & porstCOmCpbCab!NroCpb, vbExclamation
             Else
                 xImpTCb = Format(IIf(xTpoTcb = TPOTCB_VTA_IND, !ImpTCb_Vta, !ImpTCb_Cpr), FORMATO_NUM_2)
             End If
           ']
         Else
         '            frmMCpbDet.txtDato(8).Text = Format(0, FORMATO_NUM_2)
         End If
        End With
        'fin de cambio
        If porstCOmCpbDet!tpomon = "N" Then
           xTpomon = 0
        Else
           xTpomon = 1
        End If
        nImpMN = Format(porstCOmCpbDet!ImpMN, FORMATO_NUM_1)
        nImpME = Format(porstCOmCpbDet!ImpME, FORMATO_NUM_1)
        If xTpomon = TPOMON_EXT_IND Then  'And (CDec(txtImporte(0).Text) = 0 Or CDec(txtImporte(2).Text) <> CDec(txtImporte(2).Tag)) Then
          nImpMN = Format(gfRedond(CDec(nImpME) * CDec(xImpTCb), 2), FORMATO_NUM_1)
        End If
        If xTpomon = TPOMON_NAC_IND Then  'And (CDec(txtImporte(3).Text) = 0 Or CDec(txtImporte(1).Text) <> CDec(txtImporte(1).Tag)) Then
          nImpME = Format(gfRedond(CDec(nImpMN) / CDec(xImpTCb), 2), FORMATO_NUM_1)
        End If
        With porstCOCpbDet
           .AddNew
           !codemp = gsCodEmp
           !pdoano = gsAnoAct
           !mespvs = gsMesAct
           'a que diario ira
           'los formatos deben reflejar el diario
           'que mantendran en la copia
           'debe ser el mismo diario
           'porstCOmCpbDet
           !coddro = porstCOmCpbDet!coddro
           !NroCpb = sNroComprobante
           !NroIte = nNroItem
           !blqite = porstCOmCpbDet!blqite
           !codtdc = porstCOmCpbDet!codtdc
           !fehope = xDia & "/" & gsMesAct & "/" & gsAnoAct
           !codcta = porstCOmCpbDet!codcta
           !codcco = porstCOmCpbDet!codcco
           !codaux = porstCOmCpbDet!codaux
           !serdoc = porstCOmCpbDet!serdoc
           !nrodoc = porstCOmCpbDet!nrodoc
  'xFecha
           '2008/jul/21 rcs error en creacion fecha
'           !FeEDoc = xDia & "/" & gsMesAct & "/" & gsAnoAct
'           !fevdoc = xDia & "/" & gsMesAct & "/" & gsAnoAct
'           !ferdoc = xDia & "/" & gsMesAct & "/" & gsAnoAct
'           !FehOpe = xDia & "/" & gsMesAct & "/" & gsAnoAct
           
           !feedoc = xFecha
           !fevdoc = xFecha
           !ferdoc = xFecha
           !fehope = xFecha
          
           '?estamos dabo por defecto la misma fecha
           'de registro para todas las fechas
           '!FeEDoc = porstCOmCpbDet!FeEDoc
           '!fevdoc = porstCOmCpbDet!fevdoc
           '!ferdoc = porstCOmCpbDet!ferdoc
           !refdoc = porstCOmCpbDet!refdoc
           'no esta dendro de las campos a pasar
           'adicionar al select
           '!pdocpr = porstCOmCpbDet!pdocpr
           !GloIte = porstCOmCpbDet!GloIte
           !GloItex = porstCOmCpbDet!GloItex
             '!TpoCtb = IIf(IsNull(porstCOmCpbDet!TpoCtb), "", porstCOmCpbDet!TpoCtb)
           !TpoCtb = porstCOmCpbDet!TpoCtb
           !TpoPvs = porstCOmCpbDet!TpoPvs
           !tpomon = porstCOmCpbDet!tpomon
           !TpoTcb = porstCOmCpbDet!TpoTcb
        '            !ImpTCb = porstCOmCpbDet!ImpTCb
           !ImpTCb = xImpTCb
           !ImpMN = nImpMN
           !ImpME = nImpME
           !tpognr = porstCOmCpbDet!tpognr
           !indfjo_det = porstCOmCpbDet!indfjo_det
           'no esta dendro de las campos a pasar
           'adicionar al select
        '     !IndGnr_RP = porstCOmCpbDet!IndGnr_RP
           !UsrCre = gsAbvUsr
           !FyHCre = Now
        '       !usrmdf = porstCOCpbDet!usrmdf
        '       !fyhmdf = porstCOCpbDet!fyhmdf
           .Update
        End With
        porstCOmCpbDet.MoveNext
    Loop
'siguiente registro
End If
porstCOmCpbCab.MoveNext
Loop

 
Set uorstTGTCb = Nothing
Set uorstCoCta = Nothing
 
Set porstCOCpbCab = Nothing
Set porstCOCpbDet = Nothing
Set porstCoCpbAjD = Nothing
Set porstUltCoCpb = Nothing

Set porstCOmCpbCab = Nothing
Set porstCOmCpbDet = Nothing

Set uorstCOmCpbCab_ctr = Nothing
Set porstCOmCpbCab_ctr = Nothing

    
'pocnnMain.CommitTrans
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   porstCodro.Close
   'porstCoFjo.Close
   pocnnMain.Close
   
'   Set porstCoTCbMes = Nothing
'   Set porstCodro = Nothing
'   Set porstCoFjo = Nothing
   Set pocnnMain = Nothing
   
   Set porstCOmCpbCab_ok = Nothing

End Sub

Private Sub MSFlexGrid1_Click()
    ' le pasa el número de campo de tipo Boolean _
      El recordset enlazado al FlexGRid, y el control MsFlexgrid
    Call ActualizarCampo(0, porstCOmCpbCab_ok, MSFlexGrid1)
End Sub
Private Sub ActualizarCampo(nCampo As Integer, _
                            objrs As ADODB.Recordset, _
                            FlexGrid As Object, _
                            Optional character As Integer)
     
  
  With FlexGrid
    If (.MouseCol) = nCampo And .MouseRow <> 0 Then
        ' mueve al primer registro
        objrs.MoveFirst
        ' se posiciona en el registro que corresponde  la fila
        objrs.Move .row - 1
           
        ' CheckBox en false
        If .TextMatrix(.row, nCampo) = Chr(168) Then
            objrs(nCampo).Value = True ' modifica
            objrs.Update ' actualizar
            .TextMatrix(.row, nCampo) = Chr(254)
           
        ' CheckBox en true
        Else
            .TextMatrix(.row, nCampo) = Chr(168)
            objrs(nCampo).Value = False
            objrs.Update 'actualiza el recordset
               
        End If
    End If
  End With
End Sub


