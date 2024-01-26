VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fCalculoPlanilla 
   Caption         =   "Proceso de Cálculo"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "calcuplanilla.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   6780
   Begin Threed.SSCheck chkProceso 
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   9
      Top             =   2040
      Width           =   1620
      _Version        =   65536
      _ExtentX        =   2857
      _ExtentY        =   335
      _StockProps     =   78
      Caption         =   " Planilla General "
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin VB.TextBox txtTipoCambio 
      ForeColor       =   &H00000000&
      Height          =   280
      Left            =   2130
      TabIndex        =   8
      Top             =   1680
      Width           =   900
   End
   Begin VB.ComboBox cmbPeriodoAnoMes 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4245
      Width           =   1740
   End
   Begin VB.TextBox txtMesPeriodo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4695
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3720
      Width           =   1365
   End
   Begin VB.TextBox txtAnoPeriodo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3150
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3735
      Width           =   1365
   End
   Begin Threed.SSCommand cmdProceso 
      Height          =   315
      Left            =   4515
      TabIndex        =   15
      Top             =   975
      Width           =   2025
      _Version        =   65536
      _ExtentX        =   3572
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "Procesar"
   End
   Begin VB.TextBox txtPeriodo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1605
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3720
      Width           =   1365
   End
   Begin VB.ComboBox cmbPeriodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Width           =   1740
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   120
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ComboBox cmbProcesos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   195
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   315
      Width           =   5250
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   300
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin Threed.SSFrame sfmProgreso 
      Height          =   480
      Index           =   0
      Left            =   165
      TabIndex        =   10
      Top             =   2250
      Width           =   6480
      _Version        =   65536
      _ExtentX        =   11430
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   " Procesando personal : "
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      ShadowStyle     =   1
      Begin MSComctlLib.ProgressBar pgbProgreso 
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   11
         Top             =   225
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin Threed.SSFrame sfmProgreso 
      Height          =   480
      Index           =   1
      Left            =   165
      TabIndex        =   13
      Top             =   3000
      Width           =   6480
      _Version        =   65536
      _ExtentX        =   11430
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   " Procesando personal : "
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      ShadowStyle     =   1
      Begin MSComctlLib.ProgressBar pgbProgreso 
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   14
         Top             =   240
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin Threed.SSCheck chkProceso 
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   12
      Top             =   2805
      Width           =   1620
      _Version        =   65536
      _ExtentX        =   2857
      _ExtentY        =   335
      _StockProps     =   78
      Caption         =   " Planilla Netos"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin VB.Label lblconcepto 
      Caption         =   "Procesando Concepto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3720
      TabIndex        =   20
      Top             =   2760
      Width           =   2835
   End
   Begin VB.Label lblDato 
      Caption         =   "Tipo Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   2130
      TabIndex        =   7
      Top             =   1410
      Width           =   1395
   End
   Begin VB.Label lblDato 
      Caption         =   "Fecha Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Top             =   1410
      Width           =   1395
   End
   Begin VB.Label lblPeriodo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   2130
      TabIndex        =   4
      Top             =   990
      Width           =   2160
   End
   Begin VB.Label lblDato 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   765
      Width           =   1680
   End
   Begin VB.Label lblDato 
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   1680
   End
End
Attribute VB_Name = "fCalculoPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private oCalculo As Object                              ' Objeto de clase de Cálculo
Private sDomiciliado As String * 1                      ' variable de persona domiciliada
Private sMonedaPago As String                           ' Moneda de pago del personal
Private porstVariables As New ADODB.Recordset           ' Recordset de las variables
Private sTipoPeriodo As String * 1                      ' Tipo de periodo de pago, condicion personal
Private nRemImprecisa As Integer                        ' Remuneración principal imprecisa
Private sWhereEstado As String                          ' Condicion de estado personal
Private s_OptRegistro As String                         ' Instancia del formulario activo
'[
Private Sub cmbPeriodo_Click()

  cmbPeriodoAnoMes.ListIndex = cmbPeriodo.ListIndex
  
  txtPeriodo.Text = Trim(Left(cmbPeriodo.Text, 50))
  lblPeriodo.Caption = Trim(Right(cmbPeriodo.Text, 50))
  
  txtAnoPeriodo.Text = Mid(cmbPeriodoAnoMes, 1, 4)
  txtMesPeriodo.Text = Mid(cmbPeriodoAnoMes, 5, 2)
  cmdProceso.Tag = Mid(cmbPeriodoAnoMes, 7, 1)
  gdl_Procedure.EditMask "AT", mskFecha, Mid(cmbPeriodoAnoMes, 8, 10), s_MdoData_Upd, True, ""
  
  ps_Anyo = IIf(Trim(txtAnoPeriodo.Text) <> "", txtAnoPeriodo.Text, ps_Anyo)

End Sub
Private Sub cmdProceso_Click()

  Dim s_FechaHora As String, s_OldMessage As String
  Dim porstPersonal As ADODB.Recordset
  Dim porstRemunera As ADODB.Recordset
  Dim sSQL As String
  Dim nTotalRegistros As Long, nRegistroActual As Long

  Dim sConceptoModificable As String
  Dim sConceptoResultado As String
  Dim nRemAjusteIni As Double, nRemAjusteRes As Double
  Dim nRemNetoIni As Double, nRemNetoRes As Double
  Dim nVecesAjuste As Integer

  ' Realizo las validaciones de los parametrso de cálculo
  If Trim(txtPeriodo.Text) = "" Then Beep: MsgBox "Debe indicar el periodo de proceso!", vbCritical + vbOKOnly: cmbPeriodo.SetFocus: Exit Sub
  If Not gdl_Funcion.ValidaFecha(mskFecha, 1900) Then mskFecha.SetFocus: Exit Sub
  If Right(mskFecha.ClipText, 4) <> txtAnoPeriodo Then Beep: MsgBox "Fecha debe ser del periodo de Proceso", vbExclamation: mskFecha.SetFocus: Exit Sub
  If Mid(mskFecha.ClipText, 4, 2) <> txtMesPeriodo Then Beep: MsgBox "Fecha debe ser del mes de Proceso", vbExclamation: mskFecha.SetFocus: Exit Sub
  If CDec(txtTipoCambio.Text) <= 0 Then MsgBox "Tipo de cambio no puede ser menor o igual a cero; Verifique", vbInformation: txtTipoCambio.SetFocus: Exit Sub
  If Not (chkProceso(0).Value Or chkProceso(1).Value) Then MsgBox "Selecciono opción de proceso; Verifique", vbInformation: chkProceso(0).SetFocus: Exit Sub

  If Flag_RestringeSistema = "RESTRINGIR" Then
    If Valida_LicenciaUso(ps_Anyo, ps_Fecha_LimiteProc, txtMesPeriodo, txtAnoPeriodo) = False Then
      MsgBox "Calculo de planilla no puede ser procesado" & Chr(13) & "Se requiere Actualización de Componentes" & Chr(13) & "Por favor comuniquese con el personal de Sistemas.", vbInformation
      Exit Sub
    End If
  End If
  
  If MsgBox("Seguro de Procesar el Período " & lblPeriodo.Caption, vbQuestion + vbDefaultButton2 + vbYesNo) <> vbYes Then
    cmdProceso.SetFocus
    Exit Sub
  End If

  'Inactiva Botón
  cmdProceso.Enabled = False
  ' Obtengo la fecha de proceso
  s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Cálculo Información de Planilla ..."
  
  ' Obtengo el tipo de periodo
  sTipoPeriodo = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo.Text, "TP")
  sWhereEstado = IIf(sTipoPeriodo = "L", "='I'", IIf(sTipoPeriodo = "V", "='V'", "<>'I'"))
  ' Instancio el objeto de Cálculo
  Set oCalculo = CreateObject("syslink.calculo")
    
  oCalculo.sCadenaConexion = ps_StrgConnec & ps_DataBase
  oCalculo.sClasePlanilla = ps_ClsPlanilla
  oCalculo.sTipoCalculo = Right(cmbProcesos.Text, 2)
  oCalculo.sTipoProceso = cmdProceso.Tag
  oCalculo.sCodigoPeriodo = txtPeriodo.Text
  oCalculo.sDiaProceso = Left(mskFecha.ClipText, 2)
  oCalculo.sMesProceso = txtMesPeriodo.Text
  oCalculo.sAnyoProceso = txtAnoPeriodo.Text
  oCalculo.sDesAusenciaBF = s_Estado_Ina
  
  ' Obtengo si gratificación es por dias
  sSQL = "SELECT cfg.pdoano, cfg.gratixasis "
  sSQL = sSQL & "FROM plcfgempresa cfg "
  sSQL = sSQL & "WHERE cfg.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    oCalculo.sDesAusenciaBF = porstRecordset!gratixasis
  End If
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
'  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
'  Set oCalculo.oConexion = gdl_Conexion.cn_Conexion
  
  ' Selecciono el personal a procesar marcados o general
  If s_OptRegistro = "calpllgral" Then
    sSQL = "INSERT INTO rangoimpresion (proceso, valor, usrcre, fyhcre) "
    sSQL = sSQL & "SELECT '" & s_OptRegistro & "', "
    sSQL = sSQL & "psn.codpsn, '" & ps_Usuario & "', '" & s_FechaHora & "' "
    sSQL = sSQL & "FROM plpersonal psn "
    sSQL = sSQL & "INNER JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn AND asi.codpdo='" & txtPeriodo.Text & "' "
    sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    If Not (chkProceso(0).Value And chkProceso(1).Value) Then
      sSQL = sSQL & "AND psn.remuneta='" & IIf(chkProceso(0).Value, s_Estado_Ina, s_Estado_Act) & "' "
    End If
    sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
    sSQL = sSQL & "ORDER BY codpsn"
    gdl_Conexion.Execucion sSQL, Inserta
  ElseIf s_OptRegistro = "calpllpers" Then
    For nRegistroActual = 0 To o_CalculoPersona.tdbRegistro.SelBookmarks.Count - 1
      o_CalculoPersona.tdbRegistro.Bookmark = o_CalculoPersona.tdbRegistro.SelBookmarks(nRegistroActual)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, o_CalculoPersona.tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next nRegistroActual
  End If
  
  ' Inicializa variables de cálculo
  IniciaCalculo
  '[ Procesos de acuerdo a selección de cálculo
  If chkProceso(0).Value Then                   ' Planilla general
    ' Elimina los datos cálculados
    InicializaHistorico s_Estado_Ina, s_FechaHora
    ' Inserta los datos iniciales de resultado
    BuscaConceptosDefault s_Estado_Ina, s_FechaHora
    
    nTotalRegistros = 0
    nRegistroActual = 1
    pgbProgreso(0).Value = 0
    
    'Para Cada Empleado
    sSQL = "SELECT psn.*, asi.fechacese FROM plpersonal psn "
    sSQL = sSQL & "INNER JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn AND asi.codpdo='" & txtPeriodo.Text & "' "
    sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    sSQL = sSQL & "AND psn.remuneta='" & s_Estado_Ina & "' "
    sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
    sSQL = sSQL & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
    sSQL = sSQL & "WHERE proceso='" & s_OptRegistro & "' "
    sSQL = sSQL & "AND usrcre='" & ps_Usuario & "' "
    sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "') "
    sSQL = sSQL & "ORDER BY codpsn"
    Set porstPersonal = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
    If Not (porstPersonal.EOF And porstPersonal.BOF) Then
      nTotalRegistros = CLng(porstPersonal.RecordCount)
      Do While Not porstPersonal.EOF
        nRemImprecisa = CInt(porstPersonal!remimprecisa)
        sDomiciliado = IIf(porstPersonal!naciextrapsn = s_Estado_Act, "N", "D")
        sMonedaPago = porstPersonal!pagodolar
        sfmProgreso(0).Caption = " Procesando personal : " & Trim(porstPersonal!codpsn) & " - " & Trim(porstPersonal!apepaterno) & " " & Trim(porstPersonal!apematerno) & ", " & Trim(porstPersonal!nombres) & " "
        sfmProgreso(0).Refresh
        oCalculo.sCodigoEmpleado = porstPersonal!codpsn
        oCalculo.sFechaIngreso = Format(porstPersonal!fecingreso, s_FormatoFecha)
        oCalculo.sEstadoEmpleado = porstPersonal!estadopsn
        oCalculo.sFechaCese = Format(porstPersonal!fechacese, s_FormatoFecha)
        pgbProgreso(0).Value = (nRegistroActual / nTotalRegistros) * 100
        nRegistroActual = nRegistroActual + 1
        CargaVariables porstPersonal!codpsn, txtPeriodo.Text
        ProcesaFormulas porstPersonal!codpsn
        porstPersonal.MoveNext
      Loop
    End If
    porstPersonal.Close
    sfmProgreso(0).Caption = " Planilla General Finalizado "
  End If
      
  If chkProceso(1).Value Then                   ' Planilla netos
    ' Elimina los datos cálculados
    InicializaHistorico s_Estado_Act, s_FechaHora
    ' Inserta los datos iniciales de resultado
    BuscaConceptosDefault s_Estado_Act, s_FechaHora
    
    nTotalRegistros = 0
    nRegistroActual = 1
    pgbProgreso(1).Value = 0
    
    'Para Cada Empleado
    sSQL = "SELECT psn.*, asi.fechacese FROM plpersonal psn "
    sSQL = sSQL & "INNER JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn AND asi.codpdo='" & txtPeriodo.Text & "' "
    sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    sSQL = sSQL & "AND psn.remuneta='" & s_Estado_Act & "' "
    sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
    sSQL = sSQL & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
    sSQL = sSQL & "WHERE proceso='" & s_OptRegistro & "' "
    sSQL = sSQL & "AND usrcre='" & ps_Usuario & "' "
    sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "') "
    sSQL = sSQL & "ORDER BY codpsn"
    Set porstPersonal = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
    If Not (porstPersonal.EOF And porstPersonal.BOF) Then
      nTotalRegistros = CLng(porstPersonal.RecordCount)
      Do While Not porstPersonal.EOF
        
        nRemImprecisa = CInt(porstPersonal!remimprecisa)
        sDomiciliado = IIf(porstPersonal!naciextrapsn = s_Estado_Act, "N", "D")
        sMonedaPago = porstPersonal!pagodolar
        sfmProgreso(1).Caption = " Procesando personal : " & Trim(porstPersonal!codpsn) & " - " & Trim(porstPersonal!apepaterno) & " " & Trim(porstPersonal!apematerno) & ", " & Trim(porstPersonal!nombres) & " "
        sfmProgreso(1).Refresh
        oCalculo.sCodigoEmpleado = porstPersonal!codpsn
        oCalculo.sFechaIngreso = Format(porstPersonal!fecingreso, s_FormatoFecha)
        oCalculo.sEstadoEmpleado = porstPersonal!estadopsn
        oCalculo.sFechaCese = Format(porstPersonal!fechacese, s_FormatoFecha)
        pgbProgreso(1).Value = (nRegistroActual / nTotalRegistros) * 100
        nRegistroActual = nRegistroActual + 1
        CargaVariables porstPersonal!codpsn, txtPeriodo.Text
        ProcesaFormulas porstPersonal!codpsn
        
        ' Incializa valores de netos
        sConceptoModificable = porstPersonal("variacpc")
        sConceptoResultado = porstPersonal("netocpc")
        nRemAjusteIni = 0: nRemAjusteRes = 0
        nRemNetoIni = CDec(porstPersonal!imporemuneto)
        nRemNetoRes = 0
        
        ' Obtengo remuneracion neta calculada
        sSQL = "SELECT res.codcpc, res.importe_" & Choose(sMonedaPago + 1, "mn", "me") & " AS remuneracion "
        sSQL = sSQL & "FROM plresultado res "
        sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
        sSQL = sSQL & "AND res.codpsn='" & porstPersonal!codpsn & "' "
        sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
        sSQL = sSQL & "AND res.codcpc IN('" & sConceptoModificable & "', '" & sConceptoResultado & "') "
        sSQL = sSQL & "AND res.codproce='" & Right(cmbProcesos.Text, 2) & "' "
        sSQL = sSQL & "ORDER BY codcpc"
        Set porstRemunera = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
        If Not (porstRemunera.EOF And porstRemunera.BOF) Then
          nRemAjusteIni = CDec(porstRemunera!remuneracion)
          porstRemunera.MoveNext
          nRemNetoRes = CDec(porstRemunera!remuneracion)
          porstRemunera.Close
        End If
        
        ' Verifico si existen conceptos de actualización
        If (nRemAjusteIni <> 0 And nRemNetoRes <> 0) Then
          ' Verifico si importes son los requeridos
          nRemAjusteRes = Round(nRemAjusteIni + (nRemNetoRes - nRemNetoIni), 2)
          nVecesAjuste = 0
          
          ' Proceso de buscar objetivo
          Do While (Abs(CDec(nRemNetoRes) - CDec(nRemNetoIni)) <> 0# And nVecesAjuste < 3)
            ' Elimino los importes de funciones calculados
            sSQL = "DELETE res.*  "
            sSQL = sSQL & "FROM plResultado res, plconceplanilla cxp "
            sSQL = sSQL & "WHERE res.codcls = '" & ps_ClsPlanilla & "' "
            sSQL = sSQL & "And res.codpsn = '" & porstPersonal!codpsn & "' "
            sSQL = sSQL & "And res.codpdo = '" & txtPeriodo & "' "
            sSQL = sSQL & "AND res.codproce = '" & Trim(Right(cmbProcesos, 2)) & "' "
            sSQL = sSQL & "AND cxp.codcls=res.codcls "
            sSQL = sSQL & "AND cxp.codcpc=res.codcpc "
            sSQL = sSQL & "AND cxp.clasecpc='F'"
            gdl_Conexion.Execucion sSQL, Elimina
            
            ' Busco objetivo
            nRemAjusteRes = Round(nRemAjusteRes + (nRemNetoIni - nRemNetoRes), 2)
            If sMonedaPago = s_Estado_Act Then
              sSQL = "UPDATE plResultado SET importe_me= " & nRemAjusteRes & ", importe_mn= " & Round(nRemAjusteRes * CDec(txtTipoCambio.Text), 2) & " "
              sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
              sSQL = sSQL & "AND codpsn='" & porstPersonal!codpsn & "' "
              sSQL = sSQL & "AND codpdo='" & txtPeriodo & "' "
              sSQL = sSQL & "AND codcpc='" & sConceptoModificable & "' "
              sSQL = sSQL & "AND codproce='" & Right(cmbProcesos.Text, 2) & "'"
            Else
              sSQL = "UPDATE plResultado SET importe_mn= " & nRemAjusteRes & ", importe_me= " & Round(nRemAjusteRes / CDec(txtTipoCambio.Text), 2) & " "
              sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
              sSQL = sSQL & "AND codpsn='" & porstPersonal!codpsn & "' "
              sSQL = sSQL & "AND codpdo='" & txtPeriodo & "' "
              sSQL = sSQL & "AND codcpc='" & sConceptoModificable & "' "
              sSQL = sSQL & "AND codproce='" & Right(cmbProcesos.Text, 2) & "'"
            End If
            gdl_Conexion.Execucion sSQL, Modifica
  
            ' Nuevo instancia del proceso de funciones
            ProcesaFormulas porstPersonal!codpsn
            
            ' Obtengo remuneracion neta calculada
            sSQL = "SELECT res.importe_" & Choose(sMonedaPago + 1, "mn", "me") & " AS remuneracion "
            sSQL = sSQL & "FROM plresultado res "
            sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
            sSQL = sSQL & "AND res.codpsn='" & porstPersonal!codpsn & "' "
            sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
            sSQL = sSQL & "AND res.codcpc ='" & sConceptoResultado & "' "
            sSQL = sSQL & "AND res.codproce='" & Right(cmbProcesos.Text, 2) & "' "
            sSQL = sSQL & "ORDER BY codcpc"
            Set porstRemunera = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
            If Not (porstRemunera.EOF And porstRemunera.BOF) Then
              nRemNetoRes = CDec(porstRemunera!remuneracion)
              porstRemunera.Close
              ' Verifico si resultado es mayor por un centimo
              nVecesAjuste = nVecesAjuste + IIf(((CDec(nRemNetoRes) - CDec(nRemNetoIni)) = 0.01), 1, 0)
            End If
          Loop
          ' Actualizo la remuneracion de ajuste en default
          sSQL = "UPDATE plremudefa SET imporemune=" & CDec(nRemAjusteRes) & " "
          sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
          sSQL = sSQL & "AND codpsn='" & porstPersonal!codpsn & "' "
          sSQL = sSQL & "AND codcpc='" & sConceptoModificable & "'"
          gdl_Conexion.Execucion sSQL, Modifica
          If nVecesAjuste = 3 Then
            ' Actualizo la remuneracion neta ajustada
            sSQL = "UPDATE plresultado "
            sSQL = sSQL & "SET importe_" & Choose(sMonedaPago + 1, "mn", "me") & "=" & CDec(nRemNetoIni) & " "
            sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
            sSQL = sSQL & "AND codpsn='" & porstPersonal!codpsn & "' "
            sSQL = sSQL & "AND codpdo='" & txtPeriodo.Text & "' "
            sSQL = sSQL & "AND codcpc ='" & sConceptoResultado & "' "
            sSQL = sSQL & "AND codproce='" & Right(cmbProcesos.Text, 2) & "' "
            gdl_Conexion.Execucion sSQL, Modifica
          End If
        End If
        porstPersonal.MoveNext
      Loop
    End If
    porstPersonal.Close
    sfmProgreso(1).Caption = " Planilla Netos Finalizado "
  End If
  Set oCalculo = Nothing
  
  ' Actualizo los datos de resultado
  sSQL = "UPDATE pldatoresultado dxr, plresultado res, plpersonal psn, rangoimpresion rng "
  sSQL = sSQL & "SET dxr.codafp=psn.codafp, dxr.codeps=psn.codeps, dxr.regpension=psn.regpension, dxr.naciextrapsn=psn.naciextrapsn, dxr.fecingreso=psn.fecingreso, "
  ' Temporalmente el reingreso
'  sSQL = sSQL & "dxr.reingreso=psn.reingreso, dxr.codcgo=psn.codcgo, dxr.codcdt=psn.codcdt, dxr.codubica=psn.codubica, dxr.codsec=psn.codsec, dxr.estadopsn=psn.estadopsn "
  sSQL = sSQL & "dxr.codcgo=psn.codcgo, dxr.codcdt=psn.codcdt, dxr.codubica=psn.codubica, dxr.codsec=psn.codsec, dxr.estadopsn=psn.estadopsn "
  sSQL = sSQL & "WHERE dxr.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codcls=dxr.codcls "
  sSQL = sSQL & "AND res.codpdo=dxr.codpdo "
  sSQL = sSQL & "AND res.codpsn=dxr.codpsn "
  sSQL = sSQL & "AND psn.codcls=dxr.codcls "
  sSQL = sSQL & "AND psn.codpsn=dxr.codpsn "
  sSQL = sSQL & "AND res.pdoano='" & txtAnoPeriodo & "' "
  sSQL = sSQL & "AND res.pdomes='" & txtMesPeriodo & "' "
  sSQL = sSQL & "AND rng.valor=res.codpsn "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Modifica
  
  ' Actualizo las cuentas de concepto por centro de costo y seccion
  sSQL = "UPDATE plresultado res, plctacencos ctc, pldatoresultado dxr, rangoimpresion rng "
  sSQL = sSQL & "SET res.codcta_debmn=ctc.codcta_debmn, res.codcta_habmn=ctc.codcta_habmn, "
  sSQL = sSQL & "res.codcta_debme=ctc.codcta_debme, res.codcta_habme=ctc.codcta_habme "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND res.codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  sSQL = sSQL & "AND ctc.codcls=res.codcls "
  sSQL = sSQL & "AND ctc.codcpc=res.codcpc "
  sSQL = sSQL & "AND IFNULL(ctc.codafp, '')='' "
  sSQL = sSQL & "AND dxr.codcls=res.codcls "
  sSQL = sSQL & "AND dxr.codpdo=res.codpdo "
  sSQL = sSQL & "AND dxr.codpsn=res.codpsn "
  sSQL = sSQL & "AND dxr.codcco=ctc.codcco "
  sSQL = sSQL & "AND dxr.codsec=ctc.codsec "
  sSQL = sSQL & "AND rng.valor=res.codpsn "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Modifica
  ' Actualizo las cuentas de concepto por entidad de pensión
  sSQL = "UPDATE plresultado res, plctacencos ctc, pldatoresultado dxr, rangoimpresion rng "
  sSQL = sSQL & "SET res.codcta_debmn=ctc.codcta_debmn, res.codcta_habmn=ctc.codcta_habmn, "
  sSQL = sSQL & "res.codcta_debme=ctc.codcta_debme, res.codcta_habme=ctc.codcta_habme "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND res.codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  sSQL = sSQL & "AND ctc.codcls=res.codcls "
  sSQL = sSQL & "AND ctc.codcpc=res.codcpc "
  sSQL = sSQL & "AND IFNULL(ctc.codafp, '')<>'' "
  sSQL = sSQL & "AND dxr.codcls=res.codcls "
  sSQL = sSQL & "AND dxr.codpdo=res.codpdo "
  sSQL = sSQL & "AND dxr.codpsn=res.codpsn "
  sSQL = sSQL & "AND dxr.codcco=ctc.codcco "
  sSQL = sSQL & "AND dxr.codsec=ctc.codsec "
  sSQL = sSQL & "AND dxr.codafp=ctc.codafp "
  sSQL = sSQL & "AND rng.valor=res.codpsn "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Modifica
  
  ' Actualizo los periodos de cuenta corriente
  sSQL = "UPDATE plcuentacte cte, plresultado res, rangoimpresion rng SET "
  sSQL = sSQL & "cte.abono_mn=ROUND(IFNULL(IF(cte.codmon='" & s_Codmon_mn & "', cte.abono_mn, (cte.abono_me*" & CDec(txtTipoCambio.Text) & ")), 0), 2), "
  sSQL = sSQL & "cte.abono_me=ROUND(IFNULL(IF(cte.codmon='" & s_Codmon_me & "', cte.abono_me, (cte.abono_mn/" & CDec(txtTipoCambio.Text) & ")), 0), 2), "
  sSQL = sSQL & "cte.codpdocan=res.codpdo, cte.estadoctacte='" & s_Estado_Act & "' "
  sSQL = sSQL & "WHERE cte.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND cte.codpdoprv='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND cte.numcuota<>0 "
  sSQL = sSQL & "AND res.codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  sSQL = sSQL & "AND res.codcls=cte.codcls "
  sSQL = sSQL & "AND res.codpdo=cte.codpdoprv "
  sSQL = sSQL & "AND res.codpsn=cte.codpsn "
  sSQL = sSQL & "AND res.codcpc=cte.codcpc "
  sSQL = sSQL & "AND rng.valor=res.codpsn "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Modifica
  
  ' Actualizo el estado y fecha de proceso del periodo
  sSQL = "UPDATE plperiodo "
  sSQL = sSQL & "SET estadopdo='" & s_Estado_Act & "', "
  sSQL = sSQL & "fechaproceso=DATE_FORMAT('" & Format(mskFecha, s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
  sSQL = sSQL & "tipocambio=" & CDec(txtTipoCambio.Text) & " "
  sSQL = sSQL & "WHERE codpdo='" & txtPeriodo.Text & "'"
  gdl_Conexion.Execucion sSQL, Modifica
  
  ' Elimino los registros sin importes
  sSQL = "DELETE FROM plresultado "
  sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  sSQL = sSQL & "AND (importe_mn=0 AND importe_me=0) "
  sSQL = sSQL & "AND codcpc NOT IN ('2202','3000')"
  gdl_Conexion.Execucion sSQL, Modifica
  
'  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  
  pgbProgreso(0).Value = 100
  sfmProgreso(0).Caption = " Proceso Finalizado "
  pgbProgreso(1).Value = 100
  sfmProgreso(1).Caption = " Proceso Finalizado "
  ']
  
  ' Elimino el rango de proceso
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  
' Reinicializo los mensajes
  MuestraMensaje s_OldMessage
  
  MsgBox "Proceso concluyó satisfactoriamente.", vbInformation + vbOKOnly, "Sistema de Planillas"
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim sSQL As String

  lblconcepto.Caption = ""

  'Establece Posición y Titulo del Formulario
  Me.Height = 4080: Me.Width = 6900
  gdl_Procedure.CentraFormulario Me

  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
  ' Verifico que exista y Cargo el Icono del Formulario
  Me.Icon = LoadPicture()
  sSQL = gdl_Procedure.ps_PathImagen & "proceso.ico"
  If gdl_Funcion.ExisteArchivo(sSQL) Then
    Me.Icon = LoadPicture(sSQL)
  End If

  cmbProcesos.Clear
  cmbProcesos.Locked = False
  sSQL = "SELECT codcls, codproce, desproce "
  sSQL = sSQL & "FROM plproceso WHERE codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND estadoproce='" & s_Estado_Act & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    Do While Not porstRecordset.EOF
      cmbProcesos.AddItem porstRecordset!desproce & Space(200) & porstRecordset!codproce
      porstRecordset.MoveNext
    Loop
    porstRecordset.Close
    cmbProcesos.ListIndex = 0
  End If

  cmbPeriodo.Clear
  cmbPeriodoAnoMes.Clear
  cmbPeriodo.Locked = False
  sSQL = "SELECT codpdo, despdo, anopdo, mespdo, tpopdo, fechafin "
  sSQL = sSQL & "FROM plperiodo "
  sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND anopdo='" & ps_Anyo & "'"
  sSQL = sSQL & "AND estadopdo<='" & s_Estado_Act & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    Do While Not porstRecordset.EOF
      cmbPeriodo.AddItem porstRecordset!codpdo & Space(200) & porstRecordset!despdo
      cmbPeriodoAnoMes.AddItem porstRecordset!anopdo & porstRecordset!mespdo & porstRecordset!tpopdo & Format(porstRecordset!fechafin, s_FormatoFecha)
      porstRecordset.MoveNext
    Loop
    porstRecordset.Close
    cmbProcesos.ListIndex = 0
  End If
  
  ' Inicializo la fecha de proceso
  gdl_Procedure.EditMask "AT", mskFecha, "", s_MdoData_Ins, True, "##/##/####"
  gdl_Procedure.EditText "AT", txtTipoCambio, FormatNumber(0, 3), s_MdoData_Ins, False, 7, vbRightJustify
  Call cmbPeriodo_Click

End Sub
Private Sub Form_Unload(Cancel As Integer)
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub

Function CalculaExp(ByVal sPersona As String, ByVal sConcepto As String, ByVal pExp As Variant) As Double
  Dim nIndex As Integer, nIndey As Integer
  Dim wPosOri As Integer, wPosIni As Integer
  Dim wNumPar As Integer
  Dim wVal As Variant
   
  lblconcepto.Caption = "Procesando Concepto: " & sConcepto
  lblconcepto.Refresh
  
   ' *********** Calcula Conceptos *************
  For nIndex = 1 To sys_num_concpt
    Do Until InStr(1, pExp, UCase(sys_lst_concpt(nIndex))) = 0
      wVal = CalculaConceptos(Mid(sys_lst_concpt(nIndex), 2, 10), txtPeriodo.Text, Trim(Right(cmbProcesos, 2)), sPersona)
      pExp = Replace(pExp, UCase(sys_lst_concpt(nIndex)), wVal, 1, 1)
    Loop
  Next nIndex
  ' *********** Calcula Variables *************
  For nIndex = 1 To sys_num_const
    Do Until InStr(1, pExp, UCase(sys_lst_const(nIndex))) = 0
      wVal = CalculaVariables(sys_lst_const(nIndex), sPersona, txtPeriodo.Text)
      pExp = Replace(pExp, UCase(sys_lst_const(nIndex)), wVal, 1, 1)
    Loop
  Next nIndex
  ' *********** Calcula Valores *************
  For nIndex = 1 To sys_num_valores
    Do Until InStr(1, pExp, UCase(sys_lst_valores(nIndex))) = 0
      wVal = CalculaValores(sys_lst_valores(nIndex))
      pExp = Replace(pExp, UCase(sys_lst_valores(nIndex)), wVal, 1, 1)
    Loop
  Next nIndex
  ' *********** Calcula Funciones *************
  For nIndex = 1 To sys_num_func
    Do Until InStr(1, pExp, UCase(sys_lst_func(nIndex))) = 0
      wPosOri = InStr(1, pExp, UCase(sys_lst_func(nIndex)))
      wPosIni = wPosOri + Len(sys_lst_func(nIndex))
      wNumPar = 1
      nIndey = 0
      Do Until wNumPar = 0
        nIndey = nIndey + 1
        wNumPar = wNumPar + IIf(Mid(pExp, wPosIni + nIndey, 1) = "(", 1, IIf(Mid(pExp, wPosIni + nIndey, 1) = ")", -1, 0))
      Loop
      wVal = CalculaFunciones(Trim(sys_lst_func(nIndex)), Mid(pExp, wPosIni + 1, nIndey - 1), sConcepto)
      pExp = Replace(pExp, Mid(pExp, wPosOri, nIndey + Len(sys_lst_func(nIndex)) + 1), wVal, 1, 1)
    Loop
  Next nIndex
  CalculaExp = CalculaExpNum(pExp)
  
End Function
Function CalculaExpNum(pExp As Variant) As Double
  Dim o_ScripControl As Object, o_Modulo As Object
  
  ' Instancio el objeto cálculo
  Set o_ScripControl = CreateObject("ScriptControl")
  o_ScripControl.Language = "VBScript"
  ' Instancio objeto módulo.
  Set o_Modulo = o_ScripControl.Modules.Add("Module1")
  ' Agrega código al módulo.
  o_Modulo.AddCode "Private nExpresion"
  o_Modulo.AddCode "nExpresion = " & pExp
  ' Muestra la evaluación de la expresión.
  CalculaExpNum = CDec(o_Modulo.Eval("nExpresion"))
  
  Set o_Modulo = Nothing
  Set o_ScripControl = Nothing
  
End Function
Function CalculaFunciones(pFuncName As String, pFuncValue, sConcepto As String) As Double
  Dim sys_lst_parametros()
  Dim sText As String
  Dim nParametros As Integer, wPos1 As Integer
  Dim nIndex As Integer
  
  pFuncValue = pFuncValue & ","
  sText = pFuncValue
  nParametros = stNVeces(sText, ",")
  
  ReDim sys_lst_parametros(1 To nParametros)
  wPos1 = 1
  For nIndex = 1 To nParametros
    sys_lst_parametros(nIndex) = Mid(pFuncValue, wPos1, InStr(wPos1, pFuncValue, ",") - wPos1)
    wPos1 = InStr(wPos1, pFuncValue, ",") + 1
  Next nIndex
  
  Select Case UCase(pFuncName)
    Case "ACUMULADO": CalculaFunciones = FAcumulado(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)))
    Case "BASEESSALUD": CalculaFunciones = FBaseEssalud(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), CDbl(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)))
    Case "CTSTRUNCA": CalculaFunciones = FCtsTrunca
    Case "EVALUAVALORES": CalculaFunciones = FEvaluaValores(Trim(sys_lst_parametros(1)), CDbl(sys_lst_parametros(2)), Trim(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)))
    Case "GRATIFICACION": CalculaFunciones = FGratificacion(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), Trim(sys_lst_parametros(3)))
    Case "MEDIAARMONICA": CalculaFunciones = FMediaArmonica(Trim(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), CInt(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)))
    Case "PENDIENTE": CalculaFunciones = FPendiente(Trim(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)))
    Case "PROMEDIO": CalculaFunciones = FPromedio(Trim(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), CInt(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)))
    Case "RENTAQUINTA": CalculaFunciones = FQuinta(CDec(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)), Trim(sys_lst_parametros(6)), sConcepto, sDomiciliado)
    Case "SI": CalculaFunciones = FSi(sys_lst_parametros(1), Trim(sys_lst_parametros(2)), sys_lst_parametros(3), Trim(sys_lst_parametros(4)), sys_lst_parametros(5), sys_lst_parametros(6))
    Case "SUBSIDIOENFERMEDAD": CalculaFunciones = FSubsidioEnfermedad(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), Trim(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)), Trim(sys_lst_parametros(6)))
    Case "SUBSIDIOMATERNIDAD": CalculaFunciones = FSubsidioMaternidad(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), Trim(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)), Trim(sys_lst_parametros(6)))
    Case "VACATRUNCAS": CalculaFunciones = FVacaTruncas
    Case "ULTIMOMOVIMIENTO": CalculaFunciones = FUltimoMovimiento(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)))
  End Select

End Function
Function CalculaConceptos(pConcptNumber, sPeriodo As String, sProceso As String, sPersona As String) As Double
  Dim porstCalcula As New ADODB.Recordset
  Dim sSQL As String

  CalculaConceptos = 0
  
  sSQL = "SELECT importe_mn FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND codpdo='" & sPeriodo & "' "
  sSQL = sSQL & "AND codproce_pdo='" & sProceso & "' "
  sSQL = sSQL & "AND codpsn='" & sPersona & "' "
  sSQL = sSQL & "AND codcpc='" & pConcptNumber & "' "
  Set porstCalcula = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstCalcula.EOF And porstCalcula.BOF) Then
    CalculaConceptos = CDec(porstCalcula!importe_mn)
    porstCalcula.Close
  End If
  Set porstCalcula = Nothing

End Function
Function CalculaValores(pIDValor) As Double
  Dim porstCalcula As New ADODB.Recordset
  Dim sSQL As String

  CalculaValores = 0
  
  sSQL = "SELECT valor" & txtMesPeriodo & " AS valortbl "
  sSQL = sSQL & "FROM pltablabase "
  sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND pdoano='" & ps_Anyo & "' "
  sSQL = sSQL & "AND codtbl='" & Mid(pIDValor, 3) & "'"
  Set porstCalcula = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstCalcula.EOF And porstCalcula.BOF) Then
    CalculaValores = CDec(porstCalcula!valortbl)
    porstCalcula.Close
  End If
  Set porstCalcula = Nothing

End Function
Function CalculaVariables(pConstName, sPersona As String, sPeriodo As String)
 CalculaVariables = IIf(IsNull(porstVariables(pConstName)), "0", porstVariables(pConstName))
End Function
Private Sub CargaVariables(sPersona As String, sPeriodo As String)
  Dim sSQL As String
  
  sSQL = "SELECT psn.codpsn AS Codigo_Personal, psn.fecnacimiento AS Fecha_Nacimiento, IF(psn.naciextrapsn='" & s_Estado_Act & "', 'N', 'D') AS Domiciliado, "
  sSQL = sSQL & "psn.codcco AS Centro_Costo, IF(psn.sexopsn='" & s_Estado_Ina & "', 'M', 'F') AS Sexo, IF(psn.pagodolar='" & s_Estado_Act & "', 'S', 'N') AS Pago_Dolares, psn.codafp AS Codigo_AFP, "
  sSQL = sSQL & "Round(afp.factor1/100, 4) AS Factor1_AFP, Round(afp.factor2/100, 4) AS Factor2_AFP, Round(afp.factor3/100, 4) AS Factor3_AFP, Round(afp.factor4/100, 4) AS Factor4_AFP, "
  sSQL = sSQL & "IF(psn.afpmixta<>'" & s_Estado_Ina & "', 'S', 'N') AS ComisionMixta_AFP, "
  sSQL = sSQL & "psn.numdepen AS Carga_Familiar, psn.estcivilpsn AS Estado_Civil, psn.numhijo AS Numero_Hijo, IF(psn.dctojudicial='" & s_Estado_Act & "', 'S', 'N') AS Descto_Judicial, "
  sSQL = sSQL & "psn.pordsctojudi AS Porcentaje_Descto_Ju, psn.fecingreso  AS Fecha_Ingreso, IF(psn.afilsindical='" & s_Estado_Act & "', 'S', 'N') AS Sindicalizado, "
  sSQL = sSQL & "IF(psn.cgoconfianza<>'" & s_Estado_Ina & "', 'S', 'N') AS Cargo_Confianza, IF(psn.ctsdeposito='" & s_Estado_Act & "', 'S', 'N') AS Cts_Deposito, "
  sSQL = sSQL & "IF(psn.ctsdolar='" & s_Estado_Act & "', 'S', 'N') AS Cts_Dolares, IF(psn.remintegralgrati='" & s_Estado_Act & "', 'S', 'N') AS RemuIntegral_Gratifi, "
  sSQL = sSQL & "IF(psn.remintegralvaca  ='" & s_Estado_Act & "', 'S', 'N') AS RemuIntegral_Vacacio, IF(psn.remintegralcts ='" & s_Estado_Act & "', 'S', 'N') AS RemuIntegral_Cts, "
  sSQL = sSQL & "IF(psn.remimprecisa ='" & s_Estado_Act & "', 'S', 'N') AS RemuPrinci_Imprecisa, "
  sSQL = sSQL & "IF(psn.remuneta='" & s_Estado_Act & "', 'S', 'N') AS Remuneracion_Neta, psn.codeps AS Entidad_Servicios, (eps.factoreps/100) AS Factor_EPS, psn.regpension AS Regimen_Pension, "
  sSQL = sSQL & "psn.fecingregpen AS Fecha_Regimen_Pension, IF(psn.essvida='" & s_Estado_Act & "', 'S', 'N') AS Essalud_Vida, "
  sSQL = sSQL & "(CASE WHEN psn.cobsctr='" & s_Estado_Ina & "' THEN 'N' ELSE 'S' END) AS SCTR_Salud, (CASE WHEN psn.chksctrP='" & s_Estado_Ina & "' THEN 'N' ELSE 'S' END) AS SCTR_Pension, "
  sSQL = sSQL & "psn.fecbaja AS Fecha_Baja, psn.estadopsn AS Estado, asi.diatrabajo AS Dias_Trabajado, asi.dialaboral AS Dias_Laborado, asi.dialiquidacion AS Dias_Liquidar, "
  sSQL = sSQL & "asi.diaprepostnatal AS Dias_Natalidad, asi.fechacese AS Fecha_Cese, asi.horanormal AS Hora_Normal, asi.diafalta AS Faltas, "
  sSQL = sSQL & "asi.horatipo1 AS Hora_ExtraSimple, asi.horatipo2 AS Hora_ExtraDoble, asi.horatipo3 AS Hora_Especial, asi.horatipo4 AS Hora_Nocturna, asi.tardanza AS Tardanzas, "
  sSQL = sSQL & "asi.accidente AS Dias_Accidente, asi.enfermedad AS Dias_Enfermedad, asi.diavacaciones AS Dias_Vacaciones, asi.tercerturno AS Tercer_Turno, asi.diasuspension AS Dias_Suspension, "
  sSQL = sSQL & "asi.licencia AS Licencia, asi.diaferiado AS Dias_Feriado, asi.liquidavacacion AS Vacacion_Pendiente, asi.diagratificacion AS Gratifica_Trunca, asi.opcional AS Valor_Opcional, "
  sSQL = sSQL & "asi.diavacaventa AS Vacacion_Venta, asi.diamediotm AS Dias_MedioTiempo, asi.diaparcial AS Dias_TiempoParcial, asi.horamediotm AS Hora_MedioTiempo, asi.horaparcial AS Hora_TiempoParcial, "
  sSQL = sSQL & "asi.diatradesemanal AS Dias_DescansoSemanal, asi.dialibre AS Dias_Libre, "
  sSQL = sSQL & "0 AS Cts_Anyos, 0 AS Cts_Meses, 0 AS Cts_Dias, " & CDec(txtTipoCambio.Text) & " AS Tipo_Cambio, psn.jornadalaboral AS Jornada_laboral, "
  sSQL = sSQL & "IF(cte.indgratifi='" & s_Estado_Act & "', 'S', 'N') AS CtaCte_Gratifica, "
  sSQL = sSQL & "IF(cte.tpodscto='" & s_Estado_Ina & "', 'M', IF(cte.tpodscto='" & s_Estado_Act & "', 'Q', 'D')) AS CtaCte_Descuento, "
  sSQL = sSQL & "IF(psn.chk27252='" & s_Estado_Act & "', 'S', 'N') AS L27252, "
  sSQL = sSQL & "IF(asi.indvacadelanta='" & s_Estado_Act & "', 'S', 'N') AS Vacacion_Adelantada, "
  sSQL = sSQL & "asi.diavacavencida AS Vacacion_Vencida, "
  sSQL = sSQL & "asi.diapaternidad AS Dias_Paternidad, asi.diafallecefam AS Dias_FalleceFam, "
  sSQL = sSQL & "(YEAR('" & Format(mskFecha, "yyyy-mm-dd") & "')-YEAR(psn.fecnacimiento) + IF(DATE_FORMAT('" & Format(mskFecha, s_FormatoFecha) & "','%m-%d') > DATE_FORMAT(psn.fecnacimiento,'%m-%d'),0,-1)) AS Edad_Personal "
  sSQL = sSQL & "FROM plpersonal psn "
  sSQL = sSQL & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
  sSQL = sSQL & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
  sSQL = sSQL & "LEFT JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn AND asi.codpdo='" & sPeriodo & "' "
  sSQL = sSQL & "LEFT JOIN plcuentacte cte ON psn.codcls=cte.codcls AND psn.codpsn=cte.codpsn AND cte.numcuota <> 0 AND cte.codpdoprv='" & sPeriodo & "' "
  sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND psn.codpsn='" & sPersona & "'"
  Set porstVariables = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)

End Sub
Private Sub BuscaConceptosDefault(ByVal sRemuneraNeta As String, ByVal s_FechaHora As String)
  Dim sSQL As String
  
  ' Datos adicionales
  sSQL = "INSERT INTO pldatoresultado (codcls, codpdo, codpsn, codcco, codafp, codeps, regpension, naciextrapsn, fecingreso, codcgo, codcdt, codubica, codsec, fecestado, estadopsn, usrcre, fyhcre) "
  sSQL = sSQL & "SELECT psn.codcls, '" & txtPeriodo.Text & "', psn.codpsn, psn.codcco, psn.codafp, psn.codeps, psn.regpension, "
  sSQL = sSQL & "psn.naciextrapsn, psn.fecingreso, psn.codcgo, psn.codcdt, psn.codubica, psn.codsec, psn.fecestado, psn.estadopsn, "
  sSQL = sSQL & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
  sSQL = sSQL & "FROM (plpersonal psn "
  sSQL = sSQL & "INNER JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn) "
  sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
  sSQL = sSQL & "AND asi.codpdo='" & txtPeriodo.Text & "'"
  sSQL = sSQL & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
  sSQL = sSQL & "WHERE proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "')"
  gdl_Conexion.Execucion sSQL, Inserta
  
  'Remuneraciones por Default
  sSQL = "INSERT INTO plresultado (codcls, codpdo, codproce, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, usrcre, fyhcre) "
  sSQL = sSQL & "SELECT pdf.codcls, '" & txtPeriodo.Text & "', '" & Right(cmbProcesos.Text, 2) & "', pdf.codpsn, pdf.codcpc, pcp.secuencia, pdf.codmon, "
  sSQL = sSQL & "ROUND(IF(pdf.codmon='" & s_Codmon_mn & "', pdf.imporemune, (pdf.imporemune*" & CDec(txtTipoCambio.Text) & ")), 2), "
  sSQL = sSQL & "ROUND(IF(pdf.codmon='" & s_Codmon_me & "', pdf.imporemune, (pdf.imporemune/" & CDec(txtTipoCambio.Text) & ")), 2), "
  sSQL = sSQL & "NULL, Null, Null, Null, '" & txtAnoPeriodo.Text & "', '" & txtMesPeriodo.Text & "', cpc.tipocpc, cxp.impbolecpc, "
  sSQL = sSQL & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
  sSQL = sSQL & "FROM (((((plremudefa pdf "
  sSQL = sSQL & "INNER JOIN plconceproceso pcp ON pdf.codcls=pcp.codcls AND pdf.codcpc=pcp.codcpc) "
  sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON pdf.codcls=cxp.codcls AND pdf.codcpc=cxp.codcpc) "
  sSQL = sSQL & "INNER JOIN plconcepto cpc ON pdf.codcpc=cpc.codcpc) "
  sSQL = sSQL & "INNER JOIN plpersonal psn ON pdf.codcls=psn.codcls AND pdf.codpsn=psn.codpsn) "
  sSQL = sSQL & "INNER JOIN plasistencia asi ON pdf.codcls=asi.codcls AND pdf.codpsn=asi.codpsn) "
  sSQL = sSQL & "WHERE pdf.codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND pcp.codproce= '" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
  sSQL = sSQL & "AND asi.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
  sSQL = sSQL & "WHERE proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "')"
  gdl_Conexion.Execucion sSQL, Inserta

  ' Remuneraciones exepcionales
  sSQL = "INSERT INTO plresultado (codcls, codpdo, codproce, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, usrcre, fyhcre) "
  sSQL = sSQL & "SELECT pdf.codcls, '" & txtPeriodo.Text & "', '" & Right(cmbProcesos.Text, 2) & "', pdf.codpsn, pdf.codcpc, pcp.secuencia, pdf.codmon, "
  sSQL = sSQL & "ROUND(IF(pdf.codmon='" & s_Codmon_mn & "', pdf.imporemune, (pdf.imporemune*" & CDec(txtTipoCambio.Text) & ")), 2), "
  sSQL = sSQL & "ROUND(IF(pdf.codmon='" & s_Codmon_me & "', pdf.imporemune, (pdf.imporemune/" & CDec(txtTipoCambio.Text) & ")), 2), "
  sSQL = sSQL & "NULL, Null, Null, Null, '" & txtAnoPeriodo.Text & "', '" & txtMesPeriodo.Text & "', cpc.tipocpc, cxp.impbolecpc, "
  sSQL = sSQL & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
  sSQL = sSQL & "FROM (((((plremuexce pdf "
  sSQL = sSQL & "INNER JOIN plconceproceso pcp ON pdf.codcls = pcp.codcls AND pdf.codcpc=pcp.codcpc) "
  sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON pdf.codcls=cxp.codcls AND pdf.codcpc=cxp.codcpc) "
  sSQL = sSQL & "INNER JOIN plconcepto cpc ON pdf.codcpc=cpc.codcpc) "
  sSQL = sSQL & "INNER JOIN plpersonal psn ON pdf.codcls=psn.codcls AND pdf.codpsn=psn.codpsn) "
  sSQL = sSQL & "INNER JOIN plasistencia asi ON pdf.codcls=asi.codcls AND pdf.codpsn=asi.codpsn AND pdf.codpdo=asi.codpdo) "
  sSQL = sSQL & "WHERE pdf.codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND pdf.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND pcp.codproce= '" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
  sSQL = sSQL & "WHERE proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "')"
  gdl_Conexion.Execucion sSQL, Inserta

  ' Cuentas Corrientes
  sSQL = "INSERT INTO plresultado (codcls, codpdo, codproce, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, usrcre, fyhcre) "
  sSQL = sSQL & "SELECT pdf.codcls, '" & txtPeriodo.Text & "', '" & Right(cmbProcesos.Text, 2) & "', pdf.codpsn, pdf.codcpc, pcp.secuencia, pdf.codmon, "
  sSQL = sSQL & "ROUND(IFNULL(SUM(IF(pdf.codmon='" & s_Codmon_mn & "', pdf.abono_mn, (pdf.abono_me*" & CDec(txtTipoCambio.Text) & "))), 0), 2), "
  sSQL = sSQL & "ROUND(IFNULL(SUM(IF(pdf.codmon='" & s_Codmon_me & "', pdf.abono_me, (pdf.abono_mn/" & CDec(txtTipoCambio.Text) & "))), 0), 2), "
  sSQL = sSQL & "NULL, Null, Null, Null, '" & txtAnoPeriodo.Text & "', '" & txtMesPeriodo.Text & "', cpc.tipocpc, cxp.impbolecpc, "
  sSQL = sSQL & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
  sSQL = sSQL & "FROM plcuentacte pdf "
  sSQL = sSQL & "INNER JOIN plconceproceso pcp ON pdf.codcls=pcp.codcls AND pdf.codcpc=pcp.codcpc "
  sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON pdf.codcls=cxp.codcls AND pdf.codcpc=cxp.codcpc "
  sSQL = sSQL & "INNER JOIN plconcepto cpc ON pdf.codcpc=cpc.codcpc "
  sSQL = sSQL & "INNER JOIN plpersonal psn ON pdf.codcls=psn.codcls AND pdf.codpsn=psn.codpsn "
  sSQL = sSQL & "INNER JOIN plasistencia asi ON pdf.codcls=asi.codcls AND pdf.codpsn=asi.codpsn AND pdf.codpdoprv=asi.codpdo "
  sSQL = sSQL & "WHERE pdf.codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND pdf.codpdoprv='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND pcp.codproce= '" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
  sSQL = sSQL & "WHERE proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "') "
  sSQL = sSQL & "GROUP BY codcls, codpsn, codcpc, secuencia"
  gdl_Conexion.Execucion sSQL, Inserta

  ' Actualizo proceso del periodo
  sSQL = "UPDATE plresultado res, plpersonal psn, rangoimpresion rng "
  sSQL = sSQL & "SET res.codproce_pdo='" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND res.codcls=psn.codcls "
  sSQL = sSQL & "AND res.codpsn=psn.codpsn "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND res.codpsn=rng.valor "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Modifica

End Sub
Private Function stNVeces(sText As String, sCaracter As String) As Integer

Dim nPos As Integer

stNVeces = 0
For nPos = 1 To Len(sText)
  If Mid(sText, nPos, 1) = sCaracter Then
    stNVeces = stNVeces + 1
  End If
Next

End Function
Private Sub InicializaHistorico(ByVal sRemuneraNeta As String, ByVal s_FechaHora As String)
  Dim sSQL As String

  ' Elimina los datos de cálculo
  sSQL = "DELETE res.* "
  sSQL = sSQL & "FROM plresultado res, plconceproceso prc, plpersonal psn, rangoimpresion rng "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND prc.codcls=res.codcls "
  sSQL = sSQL & "AND prc.codproce='" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "AND prc.codcpc=res.codcpc "
  sSQL = sSQL & "AND psn.codcls=res.codcls "
  sSQL = sSQL & "AND psn.codpsn=res.codpsn "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND psn.estadopsn" & sWhereEstado & " "
  sSQL = sSQL & "AND res.codpsn=rng.valor "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Elimina
  
  ' Elimina los datos adicionales de cálculo
  sSQL = "DELETE res.* "
  sSQL = sSQL & "FROM pldatoresultado res, plpersonal psn, rangoimpresion rng "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpdo='" & txtPeriodo.Text & "' "
  sSQL = sSQL & "AND psn.codcls=res.codcls "
  sSQL = sSQL & "AND psn.codpsn=res.codpsn "
  sSQL = sSQL & "AND psn.remuneta='" & sRemuneraNeta & "' "
  sSQL = sSQL & "AND res.codpsn=rng.valor "
  sSQL = sSQL & "AND rng.proceso='" & s_OptRegistro & "' "
  sSQL = sSQL & "AND rng.usrcre='" & ps_Usuario & "' "
  sSQL = sSQL & "AND rng.fyhcre='" & s_FechaHora & "'"
  gdl_Conexion.Execucion sSQL, Elimina

End Sub
Private Sub ProcesaFormulas(sPersona As String)
  Dim rs As New ADODB.Recordset
  Dim sSQL As String, sFormula As String
  Dim nValorConcepto As Double, nLineas As Integer
  Dim aLinesFormula() As String
  Dim nPosDelmitador As Integer, nPosInicial As Integer
  Dim nVeces As Integer, nSecuencia As Integer
  Dim sDatoConcepto(5, 2)

  ' Obtengo concepto e información quinta categoria
  sDatoConcepto(1, 1) = "": sDatoConcepto(2, 1) = ""
  sDatoConcepto(3, 1) = 0: sDatoConcepto(4, 1) = ""
  sDatoConcepto(5, 1) = "": sDatoConcepto(1, 2) = ""
  sDatoConcepto(2, 2) = "": sDatoConcepto(3, 2) = 0
  sDatoConcepto(4, 2) = "": sDatoConcepto(5, 2) = ""
  sSQL = "SELECT cfg.codcpc5ta_ing AS cpcingreso, cxp.codcpc AS cpcdscto, cpr.secuencia, "
  sSQL = sSQL & "cpc.tipocpc, cxp.impbolecpc "
  sSQL = sSQL & "FROM plconceplanilla cxp, plconcepto cpc, plconceproceso cpr, plcfgempresa cfg "
  sSQL = sSQL & "WHERE cxp.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND cfg.pdoano='" & ps_Anyo & "' "
  sSQL = sSQL & "AND cxp.codcpc=cfg.codcpc5ta "
  sSQL = sSQL & "AND cxp.codcpc=cpc.codcpc "
  sSQL = sSQL & "AND cxp.codcls=cpr.codcls "
  sSQL = sSQL & "AND cxp.codcpc=cpr.codcpc "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "AND cpr.codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (rs.EOF And rs.BOF) Then
    sDatoConcepto(1, 2) = rs!cpcingreso: sDatoConcepto(2, 2) = rs!cpcdscto
    sDatoConcepto(3, 2) = rs!secuencia: sDatoConcepto(4, 2) = rs!tipocpc
    sDatoConcepto(5, 2) = rs!impbolecpc
  End If
  rs.Close

  ' Obtengo concepto e información devolucion quinta
  sSQL = "SELECT cxp.codcpc AS cpcingreso, cfg.codcpc5ta AS cpcdscto, cpr.secuencia, "
  sSQL = sSQL & "cpc.tipocpc, cxp.impbolecpc "
  sSQL = sSQL & "FROM plconceplanilla cxp, plconcepto cpc, plconceproceso cpr, plcfgempresa cfg "
  sSQL = sSQL & "WHERE cxp.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND cfg.pdoano='" & ps_Anyo & "' "
  sSQL = sSQL & "AND cxp.codcpc=cfg.codcpc5ta_ing "
  sSQL = sSQL & "AND cxp.codcpc=cpc.codcpc "
  sSQL = sSQL & "AND cxp.codcls=cpr.codcls "
  sSQL = sSQL & "AND cxp.codcpc=cpr.codcpc "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "AND cpr.codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  sSQL = sSQL & "AND IFNULL(cpr.formulafun, '')='' "
  Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (rs.EOF And rs.BOF) Then
    sDatoConcepto(1, 2) = rs!cpcingreso: sDatoConcepto(2, 2) = rs!cpcdscto
    sDatoConcepto(3, 2) = rs!secuencia: sDatoConcepto(4, 2) = rs!tipocpc
    sDatoConcepto(5, 2) = rs!impbolecpc
  End If
  rs.Close
  
  ' Cadenas de Texto, Recuperar Información
  sSQL = "SELECT cxp.codcpc, cpc.descpc, cpr.secuencia, cxp.clasecpc, cxp.defaultcpc, "
  sSQL = sSQL & "cxp.impbolecpc, cpr.formulafun, cxp.imagenfun, cpc.tipocpc "
  sSQL = sSQL & "FROM plconceplanilla cxp, plconcepto cpc, plconceproceso cpr "
  sSQL = sSQL & "WHERE cxp.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND cxp.codcpc=cpc.codcpc "
  sSQL = sSQL & "AND cxp.codcls=cpr.codcls "
  sSQL = sSQL & "AND cxp.codcpc=cpr.codcpc "
  sSQL = sSQL & "AND cxp.clasecpc='F' "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "AND cpr.codproce='" & Trim(Right(cmbProcesos, 2)) & "' "
  sSQL = sSQL & "AND IFNULL(cpr.formulafun, '')<>'' "
  sSQL = sSQL & "ORDER BY secuencia "
  Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
  If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
      sFormula = Trim(IIf(IsNull(rs!formulafun), "", rs!formulafun))
      nLineas = stNVeces(sFormula, ";")
      nPosDelmitador = rs.RecordCount
      
      If Mid(sFormula, 1, 1) = "@" Then
        ReDim aLinesFormula(1 To nLineas)
        For nVeces = 1 To nLineas
          nPosDelmitador = InStr(1, Trim(sFormula), ";")
          nPosInicial = InStr(1, Trim(sFormula), "=")
          aLinesFormula(nVeces) = Trim(Mid(sFormula, (nPosInicial + 1), nPosDelmitador - (nPosInicial + 1)))
          sFormula = Trim(Mid(sFormula, nPosDelmitador + 1))
        Next
          
        For nVeces = 1 To nLineas
          If InStr(1, Trim(aLinesFormula(nVeces)), "@") = 0 Then
            aLinesFormula(nVeces) = CalculaExp(sPersona, rs!codcpc, UCase(aLinesFormula(nVeces)))
            sFormula = aLinesFormula(nVeces)
          Else
            'Reemplazar las Lìneas por Valores Calculados
            For nSecuencia = 1 To nLineas
              Do While InStr(1, aLinesFormula(nVeces), "@" + Format(nSecuencia, "00"))
                aLinesFormula(nVeces) = Replace(aLinesFormula(nVeces), "@" + Format(nSecuencia, "00"), aLinesFormula(nSecuencia), 1, 1)
              Loop
            Next
            aLinesFormula(nVeces) = CalculaExp(sPersona, rs!codcpc, UCase(aLinesFormula(nVeces)))
            sFormula = aLinesFormula(nVeces)
          End If
        Next
      End If
        
      If sFormula <> "" Then
        ' Obtengo los datos de grabacion
        nValorConcepto = CalculaExp(sPersona, rs!codcpc, UCase(sFormula))
        
        sDatoConcepto(1, 1) = rs!codcpc: sDatoConcepto(2, 1) = rs!codcpc
        sDatoConcepto(3, 1) = rs!secuencia: sDatoConcepto(4, 1) = rs!tipocpc
        sDatoConcepto(5, 1) = rs!impbolecpc
        ' Verifico si concepto es de quinta y negativo
        If (sDatoConcepto(1, 1) = sDatoConcepto(2, 2) And nValorConcepto < 0) Then
          If Not (IsNull(sDatoConcepto(1, 2)) Or sDatoConcepto(1, 2) = "") Then
            sDatoConcepto(1, 1) = sDatoConcepto(1, 2): sDatoConcepto(2, 1) = sDatoConcepto(1, 2)
            sDatoConcepto(3, 1) = rs!secuencia: sDatoConcepto(4, 1) = sDatoConcepto(4, 2)
            sDatoConcepto(5, 1) = sDatoConcepto(5, 2): nValorConcepto = (nValorConcepto * -1)
          Else
            nValorConcepto = 0
          End If
        End If
        
        sSQL = "INSERT INTO plresultado "
        sSQL = sSQL & "( "
        sSQL = sSQL & "codcls, "
        sSQL = sSQL & "codpdo, "
        sSQL = sSQL & "codproce, "
        sSQL = sSQL & "codpsn, "
        sSQL = sSQL & "codcpc, "
        sSQL = sSQL & "secuencia, "
        sSQL = sSQL & "codmon, "
        sSQL = sSQL & "importe_mn, "
        sSQL = sSQL & "importe_me, "
        sSQL = sSQL & "pdoano, pdomes, "
        sSQL = sSQL & "tipocpc, impbolecpc, "
        sSQL = sSQL & "codproce_pdo, "
        sSQL = sSQL & "usrcre, "
        sSQL = sSQL & "fyhcre, "
        sSQL = sSQL & "usrmdf, "
        sSQL = sSQL & "fyhmdf "
        sSQL = sSQL & ") "
        sSQL = sSQL & "VALUES( "
        sSQL = sSQL & "'" & ps_ClsPlanilla & "', "
        sSQL = sSQL & "'" & txtPeriodo.Text & "', "
        sSQL = sSQL & "'" & Trim(Right(cmbProcesos, 2)) & "', "
        sSQL = sSQL & "'" & sPersona & "', "
        sSQL = sSQL & "'" & sDatoConcepto(1, 1) & "', "
        sSQL = sSQL & "'" & sDatoConcepto(3, 1) & "', "
        sSQL = sSQL & "'" & Choose(sMonedaPago + 1, "N", "E") & "', "
        sSQL = sSQL & "'" & Round(CDec(nValorConcepto), 2) & "', "
        sSQL = sSQL & "'" & Round(CDec(nValorConcepto) / CDec(txtTipoCambio.Text), 2) & "', "
        sSQL = sSQL & "'" & txtAnoPeriodo & "', "
        sSQL = sSQL & "'" & txtMesPeriodo & "', "
        sSQL = sSQL & "'" & sDatoConcepto(4, 1) & "', "
        sSQL = sSQL & "'" & sDatoConcepto(5, 1) & "', "
        sSQL = sSQL & "'" & Trim(Right(cmbProcesos, 2)) & "', "
        sSQL = sSQL & "'" & ps_Usuario & "', "
        sSQL = sSQL & "'" & Format(Now, s_FmtFeHoMysql_0) & "', "
        sSQL = sSQL & "NULL, NULL"
        sSQL = sSQL & ")"
        gdl_Conexion.Execucion sSQL, Inserta
      End If
      rs.MoveNext
    Loop
    rs.Close
  End If

End Sub
Private Function FAcumulado(ByVal sConcepto As String, ByVal sDelMes As String) As Double
  FAcumulado = oCalculo.Acumulado(Mid(sConcepto, 2), sDelMes)
End Function
Private Function FBaseEssalud(ByVal sBaseImponible As String, ByVal sAporteEssalud As String, ByVal nSueldoMinimo As Double, ByVal sRemuSubsidio As String) As Double
  FBaseEssalud = oCalculo.BaseEssalud(Mid(sBaseImponible, 2), Mid(sAporteEssalud, 2), nSueldoMinimo, Mid(sRemuSubsidio, 2))
End Function
Private Function FCtsTrunca() As Long
  FCtsTrunca = oCalculo.CtsTrunca
End Function
Private Function FEvaluaValores(ByVal sConcepto As String, ByVal nValorAnalizar As Double, ByVal sLimite As String, ByVal sDelMes As String) As Double
  FEvaluaValores = oCalculo.EvaluaValores(Mid(sConcepto, 2), nValorAnalizar, sLimite, sDelMes)
End Function
Private Function FGratificacion(ByVal sMesPago As String, ByVal sMesyDia As String, sProyeccion As String) As Double
  FGratificacion = oCalculo.DiasGratificacion(sMesPago, sMesyDia, sProyeccion)
End Function
Private Function FMediaArmonica(sConcepto As String, nMeses As Integer, nRepetirRegular As Integer, nRepetirImprecisa As Integer, sIncluirMes As String) As Double
  FMediaArmonica = oCalculo.Promedio(Mid(sConcepto, 2), nMeses, Choose(nRemImprecisa + 1, nRepetirRegular, nRepetirImprecisa), sIncluirMes, "D")
End Function
Private Function FPendiente(sConcepto As String, nPendiente As Integer) As Double
  FPendiente = oCalculo.Pendiente(Mid(sConcepto, 2), nPendiente)
End Function
Private Function FPromedio(sConcepto As String, nMeses As Integer, nRepetirRegular As Integer, nRepetirImprecisa As Integer, sIncluirMes As String) As Double
  FPromedio = oCalculo.Promedio(Mid(sConcepto, 2), nMeses, Choose(nRemImprecisa + 1, nRepetirRegular, nRepetirImprecisa), sIncluirMes, "M")
End Function
Private Function FQuinta(ByVal nUit As Double, ByVal nRetenido As Integer, ByVal nDividir As Integer, ByVal sGanados As String, ByVal sPendientes As String, ByVal sExtraordinarios As String, ByVal sConcepto As String, ByVal sDomiciliado As String) As Double
  FQuinta = oCalculo.RentaQuinta(nUit, nRetenido, nDividir, Mid(sGanados, 2), Mid(sPendientes, 2), Mid(sExtraordinarios, 2), sConcepto, sDomiciliado)
End Function
Private Function FSi(sCondicion1, sOperador As String, sCondicion2, sTipoComparacion As String, sVerdadero, sFalso) As Double

  ' Comparación Caracteres
  If sTipoComparacion = "C" Then
    Select Case sOperador
     Case "="
      FSi = CalculaExpNum(IIf((sCondicion1 = sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">"
      FSi = CalculaExpNum(IIf((sCondicion1 > sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<"
      FSi = CalculaExpNum(IIf((sCondicion1 < sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">="
      FSi = CalculaExpNum(IIf((sCondicion1 >= sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<="
      FSi = CalculaExpNum(IIf((sCondicion1 <= sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<>"
      FSi = CalculaExpNum(IIf((sCondicion1 <> sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "AND"
      FSi = CalculaExpNum(IIf((sCondicion1 And sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "OR"
      FSi = CalculaExpNum(IIf((sCondicion1 Or sCondicion2), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
    End Select
  End If
  
  ' Comparación Fechas
  If sTipoComparacion = "F" Then
    Select Case sOperador
     Case "="
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) = CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">"
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) > CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<"
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) < CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">="
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) >= CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<="
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) <= CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<>"
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) <> CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "AND"
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) And CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "OR"
      FSi = CalculaExpNum(IIf((CDate(sCondicion1) Or CDate(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
    End Select
  End If
  
  ' Comparación Lógico
  If sTipoComparacion = "L" Then
    Select Case sOperador
     Case "="
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) = CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">"
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) > CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<"
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) < CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">="
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) >= CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<="
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) <= CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<>"
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) <> CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "AND"
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) And CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "OR"
      FSi = CalculaExpNum(IIf((CDbl(sCondicion1) Or CDbl(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
    End Select
  End If

  ' Comparación Números
  If sTipoComparacion = "N" Then
    Select Case sOperador
     Case "="
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) = CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">"
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) > CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<"
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) < CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case ">="
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) >= CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<="
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) <= CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "<>"
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) <> CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "AND"
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) And CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
     Case "OR"
      FSi = CalculaExpNum(IIf((CalculaExpNum(sCondicion1) Or CalculaExpNum(sCondicion2)), CalculaExpNum(sVerdadero), CalculaExpNum(sFalso)))
    End Select
  End If

End Function
Private Function FSubsidioEnfermedad(ByVal sRemuneraFija As String, ByVal sRemuneraVariable1 As String, ByVal sRemuneraVariable2 As String, ByVal sRemuneraVariable3 As String, ByVal sRemuneraVariable4 As String, ByVal sRemuneraVariable5 As String) As Double
  
  ' Valido dias de subsidio enfermedad
  If (CLng(porstVariables("Dias_Enfermedad")) <= 0) Then GoTo Finaliza
  FSubsidioEnfermedad = oCalculo.SubsidioEnfermedad(Mid(sRemuneraFija, 2), Mid(sRemuneraVariable1, 2), Mid(sRemuneraVariable2, 2), Mid(sRemuneraVariable3, 2), Mid(sRemuneraVariable4, 2), Mid(sRemuneraVariable5, 2), porstVariables("Dias_Accidente"))
Finaliza:

End Function
Private Function FSubsidioMaternidad(ByVal sRemuneraFija As String, ByVal sRemuneraVariable1 As String, ByVal sRemuneraVariable2 As String, ByVal sRemuneraVariable3 As String, ByVal sRemuneraVariable4 As String, ByVal sRemuneraVariable5 As String) As Double
 
  ' Valido dias de subsidio maternidad, sexo Femenino
  If Not (CLng(porstVariables("Dias_Natalidad")) > 0 And Trim(porstVariables("Sexo")) = "F") Then GoTo Finaliza
  FSubsidioMaternidad = oCalculo.SubsidioMaternidad(Mid(sRemuneraFija, 2), Mid(sRemuneraVariable1, 2), Mid(sRemuneraVariable2, 2), Mid(sRemuneraVariable3, 2), Mid(sRemuneraVariable4, 2), Mid(sRemuneraVariable5, 2))
Finaliza:

End Function
Private Function FUltimoMovimiento(ByVal sConcepto As String, ByVal sTblResultado As String, ByVal nMesAnterior As Integer, ByVal sTipoProceso As String) As Double
  sTblResultado = "pl" & IIf(sTblResultado = "C", "cts", "") & "resultado"
  FUltimoMovimiento = oCalculo.UltimoMovimiento(Mid(sConcepto, 2), sTblResultado, nMesAnterior, sTipoProceso)
End Function
Private Function FVacaTruncas() As Long
  FVacaTruncas = oCalculo.VacacionTrunca
End Function

Private Sub mskFecha_GotFocus()
  gdl_Procedure.MarcaGet mskFecha
End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecha_LostFocus()

  If mskFecha.Tag <> mskFecha Then
    ' Obtengo el tipo de cambio de acuerdo ala fecha
    txtTipoCambio.Text = FormatNumber(gdl_Funcion.Tipocambio(ps_StrgConnec & ps_DaBasCon, mskFecha, "V"), 3)
    mskFecha.Tag = Format(mskFecha, s_FormatoFecha)
    If CDec(txtTipoCambio.Text) <= 0 Then MsgBox "No se ha ingresado tipo de cambio para la fecha; Verifique", vbInformation: mskFecha.SetFocus: Exit Sub
  End If

End Sub
Private Sub mskFecha_Validate(Cancel As Boolean)
  If mskFecha.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecha, 1900) Then mskFecha.SetFocus: Cancel = True: Exit Sub
  End If
  If Right(mskFecha.ClipText, 4) <> txtAnoPeriodo Then Beep: MsgBox "Fecha debe ser del periodo a Procesar", vbExclamation: mskFecha.SetFocus: Cancel = True: Exit Sub
  If Mid(mskFecha.ClipText, 4, 2) <> txtMesPeriodo Then Beep: MsgBox "Fecha debe ser del mes a Procesar", vbExclamation: mskFecha.SetFocus: Cancel = True: Exit Sub
End Sub
Private Sub txtTipoCambio_GotFocus()
  gdl_Procedure.MarcaGet txtTipoCambio
End Sub
Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoCambio_Validate(Cancel As Boolean)
  txtTipoCambio.Text = IIf(Not IsNumeric(txtTipoCambio.Text), 0, txtTipoCambio.Text)
  If CDec(txtTipoCambio.Text) <= 0 Then MsgBox "Tipo de cambio no puede ser menor o igual a cero; Verifique", vbInformation: txtTipoCambio.SetFocus: Cancel = True: Exit Sub
  txtTipoCambio.Text = FormatNumber(txtTipoCambio.Text, 3)
End Sub
