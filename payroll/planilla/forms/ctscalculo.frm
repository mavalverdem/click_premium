VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fCtsCalculo 
   Caption         =   "Proceso de Cálculo"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "ctscalculo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   6690
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2235
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1515
      Width           =   1860
   End
   Begin VB.TextBox txtTipoCambio 
      ForeColor       =   &H00000000&
      Height          =   280
      Left            =   2235
      TabIndex        =   7
      Top             =   2205
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
      TabIndex        =   14
      Top             =   4020
      Width           =   1740
   End
   Begin VB.TextBox txtMesPeriodo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4695
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3405
      Width           =   1365
   End
   Begin VB.TextBox txtAnoPeriodo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   3150
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3420
      Width           =   1365
   End
   Begin Threed.SSCommand cmdProceso 
      Height          =   315
      Left            =   4515
      TabIndex        =   10
      Top             =   1500
      Width           =   2025
      _Version        =   65536
      _ExtentX        =   3572
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "Procesar"
      Picture         =   "ctscalculo.frx":000C
   End
   Begin VB.TextBox txtPeriodo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1605
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3405
      Width           =   1365
   End
   Begin VB.ComboBox cmbEjercicio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1515
      Width           =   1740
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   120
      Top             =   3375
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   5250
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   2205
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
      TabIndex        =   8
      Top             =   2640
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
         TabIndex        =   9
         Top             =   225
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6690
      _Version        =   65536
      _ExtentX        =   11800
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Width           =   5490
      End
   End
   Begin VB.Label lblDato 
      Caption         =   "Mes :"
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
      Index           =   4
      Left            =   2235
      TabIndex        =   18
      Top             =   1290
      Width           =   1680
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
      Left            =   2235
      TabIndex        =   6
      Top             =   1935
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
      Left            =   240
      TabIndex        =   4
      Top             =   1935
      Width           =   1395
   End
   Begin VB.Label lblDato 
      Caption         =   "Ejercicio :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1290
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
      Top             =   615
      Width           =   1680
   End
End
Attribute VB_Name = "fCtsCalculo"
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
Private sGratixAsis As String * 1                       ' Dias de gratificación  por asistencia
Private nContador As Integer                            ' Indice de rango
'[
Private Sub cmbPeriodo_Click()
  Dim sFecha As String
  
  ' Obtengo al fecha final del mes
  sFecha = gdl_Funcion.NumeroDiasMes(Left(Trim(cmbPeriodo), 2), Trim(cmbejercicio)) & "/" & Left(Trim(cmbPeriodo), 2) & "/" & Trim(cmbejercicio)
  gdl_Procedure.EditMask "AT", mskFecha, sFecha, s_MdoData_Upd, True, "##/##/####"
  txtPeriodo.Text = Trim(cmbejercicio) & Trim(Left(cmbPeriodo.Text, 2))
  
  txtMesPeriodo.Text = Left(cmbPeriodo, 2)

End Sub

Private Sub cmdProceso_Click()
  Dim sSQL As String
  Dim nTotalRegistros As Long, nRegistroActual As Long, nRegistros As Long

  ' Realizo las validaciones de los parametrso de cálculo
  If Trim(cmbejercicio) = "" Then Beep: MsgBox "Debe indicar el ejercicio de proceso!", vbCritical + vbOKOnly: cmbejercicio.SetFocus: Exit Sub
  If Trim(cmbPeriodo) = "" Then Beep: MsgBox "Debe indicar el periodo de proceso!", vbCritical + vbOKOnly: cmbPeriodo.SetFocus: Exit Sub
  If Not gdl_Funcion.ValidaFecha(mskFecha, 1900) Then mskFecha.SetFocus: Exit Sub
  If Right(mskFecha.ClipText, 4) <> Trim(cmbejercicio) Then Beep: MsgBox "Fecha debe ser del periodo de Proceso", vbExclamation: mskFecha.SetFocus: Exit Sub
  If Mid(mskFecha.ClipText, 3, 2) <> Left(Trim(cmbPeriodo), 2) Then Beep: MsgBox "Fecha debe ser del mes de Proceso", vbExclamation: mskFecha.SetFocus: Exit Sub
  If CDec(txtTipoCambio.Text) <= 0 Then MsgBox "Tipo de cambio no puede ser menor o igual a cero; Verifique", vbInformation: txtTipoCambio.SetFocus: Exit Sub
  
  If MsgBox("Seguro de Procesar el Período " & Trim(cmbejercicio) & "-" & Trim(cmbPeriodo), vbQuestion + vbDefaultButton2 + vbYesNo) <> vbYes Then
    cmdProceso.SetFocus
    Exit Sub
  End If

  'Inactiva Botón
  cmdProceso.Enabled = False
  
  ' Obtengo si BBSS es por dias
  sGratixAsis = s_Estado_Ina
  sSQL = "SELECT cfg.pdoano, cfg.gratixasis "
  sSQL = sSQL & "FROM plcfgempresa cfg "
  sSQL = sSQL & "WHERE cfg.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    sGratixAsis = porstRecordset!gratixasis
  End If
  
  ' Obtengo el tipo de periodo
  sTipoPeriodo = "N"
  ' Instancio el objeto de Cálculo
  Set oCalculo = CreateObject("syslink.calculo")
    
  oCalculo.sCadenaConexion = ps_StrgConnec & ps_DataBase
  oCalculo.sClasePlanilla = ps_ClsPlanilla
  oCalculo.sTipoCalculo = Right(cmbProcesos.Text, 2)
  oCalculo.sTipoProceso = sTipoPeriodo
  oCalculo.sCodigoPeriodo = txtPeriodo.Text
  oCalculo.sDiaProceso = Left(mskFecha.ClipText, 2)
  oCalculo.sMesProceso = Left(Trim(cmbPeriodo), 2)
  oCalculo.sAnyoProceso = Trim(cmbejercicio)
  oCalculo.sDesAusenciaBF = sGratixAsis
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
'  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
'  Set oCalculo.oConexion = gdl_Conexion.cn_Conexion
  
  ' Inicializa variables de cálculo
  IniciaCalculo
  '[ Procesos de acuerdo a selección de cálculo
  nTotalRegistros = 0
  nRegistroActual = 1
  pgbProgreso(0).Value = 0
  
  ' Obtengo el periodo y sub periodo de cts
  sSQL = "SELECT DISTINCTROW sub.pdocts, sub.subcts "
  sSQL = sSQL & "FROM plctsperiodosub sub "
  sSQL = sSQL & "INNER JOIN plctsperiodo pdo ON sub.codcls=pdo.codcls AND sub.pdocts=pdo.pdocts "
  sSQL = sSQL & "WHERE sub.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND sub.estadosub<>'" & s_Estado_Blq & "' "
  sSQL = sSQL & "AND sub.pdoano='" & Trim(cmbejercicio) & "' "
  sSQL = sSQL & "AND sub.pdomes='" & Left(Trim(cmbPeriodo), 2) & "' "
  sSQL = sSQL & "ORDER BY pdocts, subcts"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    cmbejercicio.Tag = Trim(porstRecordset!pdocts)
    cmbPeriodo.Tag = Trim(porstRecordset!subcts)
    porstRecordset.Close
  End If
  
  'Para Cada Empleado
  nTotalRegistros = o_PvsComxTieSer.tdbRegistro.SelBookmarks.Count
  For nContador = 0 To o_PvsComxTieSer.tdbRegistro.SelBookmarks.Count - 1
    o_PvsComxTieSer.tdbRegistro.Bookmark = o_PvsComxTieSer.tdbRegistro.SelBookmarks(nContador)
    ' Actualizo el registro del personal
    sfmProgreso(0).Caption = " Procesando personal : " & Trim(o_PvsComxTieSer.dcaRegistro.Recordset!codpsn) & " - " & Trim(o_PvsComxTieSer.dcaRegistro.Recordset!nombrepsn) & " "
    sfmProgreso(0).Refresh
    ' Verifico si existe información para Cálculo
    nRegistros = 0
    sSQL = "SELECT COUNT(*) AS Registros "
    sSQL = sSQL & "FROM plctsmovimiento mov "
    sSQL = sSQL & "INNER JOIN plctsperiodo pdo ON mov.codcls=pdo.codcls AND mov.pdocts=pdo.pdocts "
    sSQL = sSQL & "WHERE mov.codcls='" & ps_ClsPlanilla & "' "
    sSQL = sSQL & "AND mov.codpsn='" & o_PvsComxTieSer.dcaRegistro.Recordset!codpsn & "' "
    sSQL = sSQL & "AND mov.estadomov<>'" & s_Estado_Blq & "' "
    sSQL = sSQL & "AND mov.pdomes='" & Left(Trim(cmbPeriodo), 2) & "' "
    sSQL = sSQL & "AND mov.pdoano='" & Trim(cmbejercicio) & "' "
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
    If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      nRegistros = CLng(porstRecordset!registros)
      porstRecordset.Close
    End If
    ' Si existe periodos provisionados
    If nRegistros > 0 Then
      ' Elimina los datos cálculados
      sSQL = "DELETE res.* "
      sSQL = sSQL & "FROM plctsresultado res "
      sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      sSQL = sSQL & "AND res.pdoano='" & Trim(cmbejercicio) & "' "
      sSQL = sSQL & "AND res.pdomes='" & Left(Trim(cmbPeriodo), 2) & "' "
      sSQL = sSQL & "AND res.codpsn='" & o_PvsComxTieSer.dcaRegistro.Recordset!codpsn & "' "
      gdl_Conexion.Execucion sSQL, Elimina
      ' Inserta los datos iniciales de resultado
      ConceptosIniciales o_PvsComxTieSer.dcaRegistro.Recordset!codpsn
      
      ' Variables del objeto de Cálculo
      sDomiciliado = IIf(o_PvsComxTieSer.dcaRegistro.Recordset!naciextrapsn = s_Estado_Act, "N", "D")
      nRemImprecisa = CInt(o_PvsComxTieSer.dcaRegistro.Recordset!remimprecisa)
      sMonedaPago = o_PvsComxTieSer.dcaRegistro.Recordset!pagodolar
      oCalculo.sCodigoEmpleado = o_PvsComxTieSer.dcaRegistro.Recordset!codpsn
      oCalculo.sFechaIngreso = Format(o_PvsComxTieSer.dcaRegistro.Recordset!fecingreso, s_FormatoFecha)
      oCalculo.sEstadoEmpleado = o_PvsComxTieSer.dcaRegistro.Recordset!estadopsn
      oCalculo.sFechaCese = Format(o_PvsComxTieSer.dcaRegistro.Recordset!fecbaja, s_FormatoFecha)
      CargaVariables o_PvsComxTieSer.dcaRegistro.Recordset!codpsn, txtPeriodo.Text
      ProcesaFormulas o_PvsComxTieSer.dcaRegistro.Recordset!codpsn
      
      ' Actualizo movimientos de cts
      sSQL = "UPDATE plctsperiodo pdo, plctsperiodosub sub, plctsmovimiento mov  "
      sSQL = sSQL & "SET pdo.estadocts='" & s_Estado_Act & "', sub.estadosub='" & s_Estado_Act & "', "
      sSQL = sSQL & "mov.tipocambio=" & CDec(txtTipoCambio.Text) & ", "
      sSQL = sSQL & "mov.estadomov='" & s_Estado_Act & "' "
      sSQL = sSQL & "WHERE pdo.codcls='" & ps_ClsPlanilla & "' "
      sSQL = sSQL & "AND pdo.pdocts='" & Trim(cmbejercicio.Tag) & "' "
      sSQL = sSQL & "AND sub.codcls=pdo.codcls "
      sSQL = sSQL & "AND sub.pdocts=pdo.pdocts "
      sSQL = sSQL & "AND sub.subcts='" & Trim(cmbPeriodo.Tag) & "' "
      sSQL = sSQL & "AND mov.codcls=sub.codcls "
      sSQL = sSQL & "AND mov.pdocts=sub.pdocts "
      sSQL = sSQL & "AND mov.subcts=sub.subcts "
      sSQL = sSQL & "AND mov.codpsn='" & Trim(o_PvsComxTieSer.dcaRegistro.Recordset!codpsn) & "' "
      gdl_Conexion.Execucion sSQL, Modifica
    End If
    pgbProgreso(0).Value = (nRegistroActual / nTotalRegistros) * 100
    nRegistroActual = nRegistroActual + 1
  Next nContador
  sfmProgreso(0).Caption = " " & lblTitle & " Finalizado "
  Set oCalculo = Nothing
  
  ' Obtengo el periodo de planilla general
  sSQL = "SELECT codpdo "
  sSQL = sSQL & "FROM plperiodo "
  sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND anopdo='" & Trim(cmbejercicio.Text) & "' "
  sSQL = sSQL & "AND mespdo='" & Left(Trim(cmbPeriodo.Text), 2) & "' "
  sSQL = sSQL & "AND tpopdo='" & sTipoPeriodo & "' "
  sSQL = sSQL & "AND estadopdo IN ('" & s_Estado_Act & "', '" & s_Estado_Blq & "')"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    txtPeriodo.Tag = Trim(porstRecordset!codpdo)
    porstRecordset.Close
  End If
  
  ' Actualizo las cuentas de concepto por centro de costo
  sSQL = "UPDATE plctsresultado res, plctacencos ctc, pldatoresultado dxr, plperiodo pdo "
  sSQL = sSQL & "SET res.codcta_debmn=ctc.codcta_debmn, res.codcta_habmn=ctc.codcta_habmn, "
  sSQL = sSQL & "res.codcta_debme=ctc.codcta_debme, res.codcta_habme=ctc.codcta_habme "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.pdocts='" & Trim(cmbejercicio.Tag) & "' "
  sSQL = sSQL & "AND res.subcts='" & Trim(cmbPeriodo.Tag) & "' "
  sSQL = sSQL & "AND ctc.codcls=res.codcls "
  sSQL = sSQL & "AND ctc.codcpc=res.codcpc "
  sSQL = sSQL & "AND IFNULL(ctc.codafp, '')='' "
  sSQL = sSQL & "AND dxr.codcls=res.codcls "
  sSQL = sSQL & "AND dxr.codpsn=res.codpsn "
  sSQL = sSQL & "AND dxr.codcco=ctc.codcco "
  sSQL = sSQL & "AND dxr.codsec=ctc.codsec "
  sSQL = sSQL & "AND pdo.codcls=dxr.codcls "
  sSQL = sSQL & "AND pdo.codpdo=dxr.codpdo "
  sSQL = sSQL & "AND CONCAT(pdo.anopdo, pdo.mespdo)='" & Trim(cmbejercicio.Text) & Left(Trim(cmbPeriodo.Text), 2) & "' "
  sSQL = sSQL & "AND pdo.tpopdo='" & sTipoPeriodo & "' "
  sSQL = sSQL & "AND pdo.estadopdo IN ('" & s_Estado_Act & "', '" & s_Estado_Blq & "')"
  gdl_Conexion.Execucion sSQL, Modifica
  ' Actualizo las cuentas de concepto por entidad de pensión
  sSQL = "UPDATE plctsresultado res, plctacencos ctc, pldatoresultado dxr, plperiodo pdo "
  sSQL = sSQL & "SET res.codcta_debmn=ctc.codcta_debmn, res.codcta_habmn=ctc.codcta_habmn, "
  sSQL = sSQL & "res.codcta_debme=ctc.codcta_debme, res.codcta_habme=ctc.codcta_habme "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.pdocts='" & Trim(cmbejercicio.Tag) & "' "
  sSQL = sSQL & "AND res.subcts='" & Trim(cmbPeriodo.Tag) & "' "
  sSQL = sSQL & "AND ctc.codcls=res.codcls "
  sSQL = sSQL & "AND ctc.codcpc=res.codcpc "
  sSQL = sSQL & "AND IFNULL(ctc.codafp, '')<>'' "
  sSQL = sSQL & "AND dxr.codcls=res.codcls "
  sSQL = sSQL & "AND dxr.codpsn=res.codpsn "
  sSQL = sSQL & "AND dxr.codcco=ctc.codcco "
  sSQL = sSQL & "AND dxr.codsec=ctc.codsec "
  sSQL = sSQL & "AND dxr.codafp=ctc.codafp "
  sSQL = sSQL & "AND pdo.codcls=dxr.codcls "
  sSQL = sSQL & "AND pdo.codpdo=dxr.codpdo "
  sSQL = sSQL & "AND CONCAT(pdo.anopdo, pdo.mespdo)='" & Trim(cmbejercicio.Text) & Left(Trim(cmbPeriodo.Text), 2) & "' "
  sSQL = sSQL & "AND pdo.tpopdo='" & sTipoPeriodo & "' "
  sSQL = sSQL & "AND pdo.estadopdo IN ('" & s_Estado_Act & "', '" & s_Estado_Blq & "')"
  gdl_Conexion.Execucion sSQL, Modifica
  
  ' Elimino los registros sin importes
  sSQL = "DELETE FROM plctsresultado "
  sSQL = sSQL & "WHERE codcls = '" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND pdocts='" & Trim(cmbejercicio.Tag) & "' "
  sSQL = sSQL & "AND subcts='" & Trim(cmbPeriodo.Tag) & "' "
  sSQL = sSQL & "AND (importe_mn=0 AND importe_me=0)"
  gdl_Conexion.Execucion sSQL, Modifica
  
'  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  
  pgbProgreso(0).Value = 100
  sfmProgreso(0).Caption = " Proceso Finalizado "
  ']
  MsgBox "Proceso concluyó satisfactoriamente.", vbInformation + vbOKOnly, "Sistema de Planillas"
  'Activo el botón de proceso
  cmdProceso.Enabled = True
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
Dim nTop As Integer
Dim nLeft As Integer
Dim nHeight As Integer
Dim nWidth As Integer
Dim rs As New ADODB.Recordset
Dim sSQL As String

nHeight = 3700
nLeft = 3000
nTop = 1400
nWidth = 6840
' Verifico que exista y Cargo el Icono del Formulario
Me.Icon = LoadPicture()
sSQL = gdl_Procedure.ps_PathImagen & "proceso.ico"
If gdl_Funcion.ExisteArchivo(sSQL) Then
  Me.Icon = LoadPicture(sSQL)
End If

With Me
    .Height = nHeight
    .Left = nLeft
    .Top = nTop
    .Width = nWidth
End With
lblTitle = "Cálculo Compensación por Tiempo de Servicio"

cmbProcesos.Clear
cmbProcesos.Locked = False
sSQL = "SELECT codcls, codproce, desproce FROM plproceso WHERE estadoproce='" & s_Estado_Act & "' AND codcls = '" & ps_ClsPlanilla & "'"
Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
  Do While Not rs.EOF
    cmbProcesos.AddItem rs("desproce") & Space(200) & rs("codproce")
    rs.MoveNext
  Loop
  rs.Close
  cmbProcesos.ListIndex = 0
End If

' Configuro los listados, datos adicionales
For nContador = (Val(ps_Anyo) - 10) To (Val(ps_Anyo) + 10): cmbejercicio.AddItem Format(nContador, "0000"): Next nContador
For nContador = 1 To 12: cmbPeriodo.AddItem Choose(nContador, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next nContador

' Inicializo la fecha de proceso
gdl_Procedure.EditMask "AT", mskFecha, "", s_MdoData_Ins, True, "##/##/####"
gdl_Procedure.EditText "AT", txtTipoCambio, FormatNumber(0, 3), s_MdoData_Ins, False, 7, vbRightJustify
gdl_Procedure.EditCombo "PK", cmbejercicio, 10, s_MdoData_Ins, True

End Sub
Private Sub Form_Unload(Cancel As Integer)
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub

Function CalculaExp(ByVal sPersona As String, ByVal sConcepto As String, ByVal pExp As Variant) As Double
Dim i As Integer
Dim j As Integer
Dim wPosOri As Integer
Dim wPosIni As Integer
Dim wNumPar As Integer
Dim wVal As Variant
  
'*********** Calcula Conceptos *************
  For i = 1 To sys_num_concpt
    Do Until InStr(1, pExp, UCase(sys_lst_concpt(i))) = 0
      wVal = CalculaConceptos(Mid(sys_lst_concpt(i), 2, 10), cmbejercicio.Tag, cmbPeriodo.Tag, sPersona)
      pExp = Replace(pExp, UCase(sys_lst_concpt(i)), wVal, 1, 1)
    Loop
  Next
'*********** Calcula Variables *************
  For i = 1 To sys_num_const
    Do Until InStr(1, pExp, UCase(sys_lst_const(i))) = 0
      wVal = CalculaVariables(sys_lst_const(i), sPersona, txtPeriodo.Text)
      pExp = Replace(pExp, UCase(sys_lst_const(i)), wVal, 1, 1)
    Loop
  Next
'*********** Calcula Valores *************
  For i = 1 To sys_num_valores
    Do Until InStr(1, pExp, UCase(sys_lst_valores(i))) = 0
      wVal = CalculaValores(sys_lst_valores(i))
      pExp = Replace(pExp, UCase(sys_lst_valores(i)), wVal, 1, 1)
    Loop
  Next
'*********** Calcula Funciones *************
  For i = 1 To sys_num_func
    Do Until InStr(1, pExp, UCase(sys_lst_func(i))) = 0
      wPosOri = InStr(1, pExp, UCase(sys_lst_func(i)))
      wPosIni = wPosOri + Len(sys_lst_func(i))
      wNumPar = 1
      j = 0
      Do Until wNumPar = 0
        j = j + 1
        wNumPar = wNumPar + IIf(Mid(pExp, wPosIni + j, 1) = "(", 1, IIf(Mid(pExp, wPosIni + j, 1) = ")", -1, 0))
      Loop
      wVal = CalculaFunciones(Trim(sys_lst_func(i)), Mid(pExp, wPosIni + 1, j - 1), sConcepto)
      pExp = Replace(pExp, Mid(pExp, wPosOri, j + Len(sys_lst_func(i)) + 1), wVal, 1, 1)
    Loop
  Next
  CalculaExp = CalculaExpNum(pExp)
End Function

Function CalculaExpNum(pExp As Variant) As Double
'On Error GoTo Err
   ' Crea algunas variables.
   Dim sc, m
   Set sc = CreateObject("ScriptControl")
   sc.Language = "VBScript"
   ' Agrega un módulo.
   Set m = sc.Modules.Add("Module1")
   ' Agrega código al módulo.
   m.AddCode "Private x"
   m.AddCode "x = " & pExp
   ' Muestra la evaluación de la expresión.
  CalculaExpNum = m.Eval("x")
'Err:
'  txtControl.Text = Err.Number
'  Resume Next
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
    Case "EVALUAVALORES": CalculaFunciones = FEvaluaValores(Trim(sys_lst_parametros(1)), CDbl(sys_lst_parametros(2)), Trim(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)))
    Case "CTSTRUNCA": CalculaFunciones = FCtsTrunca
    Case "GRATIFICACION": CalculaFunciones = FGratificacion(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), Trim(sys_lst_parametros(3)))
    Case "MEDIAARMONICA": CalculaFunciones = FMediaArmonica(Trim(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), CInt(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)))
    Case "PENDIENTE": CalculaFunciones = FPendiente(Trim(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)))
    Case "PROMEDIO": CalculaFunciones = FPromedio(Trim(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), CInt(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)))
    Case "RENTAQUINTA": CalculaFunciones = FQuinta(CDec(sys_lst_parametros(1)), CInt(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)), Trim(sys_lst_parametros(5)), Trim(sys_lst_parametros(6)), sConcepto, sDomiciliado)
    Case "SI": CalculaFunciones = FSi(sys_lst_parametros(1), sys_lst_parametros(2), sys_lst_parametros(3), sys_lst_parametros(4), sys_lst_parametros(5), sys_lst_parametros(6))
    Case "VACATRUNCAS": CalculaFunciones = FVacaTruncas
    Case "ULTIMOMOVIMIENTO": CalculaFunciones = FUltimoMovimiento(Trim(sys_lst_parametros(1)), Trim(sys_lst_parametros(2)), CInt(sys_lst_parametros(3)), Trim(sys_lst_parametros(4)))
  End Select

End Function
Function CalculaVariables(pConstName, sPersona As String, sPeriodo As String)
  CalculaVariables = IIf(IsNull(porstVariables(pConstName)), "0", porstVariables(pConstName))
End Function

Function CalculaConceptos(pConcptNumber, sPeriodo As String, sProceso As String, sPersona As String) As Double
Dim porstConcepto As New ADODB.Recordset
Dim sSQL As String

CalculaConceptos = 0

sSQL = "SELECT importe_mn FROM plctsresultado "
sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
sSQL = sSQL & "AND pdocts='" & sPeriodo & "' "
sSQL = sSQL & "AND subcts='" & sProceso & "' "
sSQL = sSQL & "AND codpsn='" & sPersona & "' "
sSQL = sSQL & "AND codcpc='" & pConcptNumber & "' "
Set porstConcepto = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)
If Not (porstConcepto.EOF And porstConcepto.BOF) Then
  CalculaConceptos = CDec(porstConcepto("importe_mn"))
  porstConcepto.Close
End If
Set porstConcepto = Nothing

End Function
Function CalculaValores(pIDValor) As Double
Dim rs As New ADODB.Recordset
Dim sSQL As String

CalculaValores = 0

sSQL = "SELECT valor" & txtMesPeriodo & " AS valortbl "
sSQL = sSQL & "FROM pltablabase "
sSQL = sSQL & "WHERE codcls='" & ps_ClsPlanilla & "' "
sSQL = sSQL & "AND pdoano='" & ps_Anyo & "' "
sSQL = sSQL & "AND codtbl='" & Mid(pIDValor, 3) & "'"


Set rs = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, sSQL)
If Not (rs.EOF And rs.BOF) Then
  CalculaValores = rs!valortbl
  rs.Close
End If

End Function
Private Sub ConceptosIniciales(ByVal sPersonal As String)
  Dim sSQL As String

  ' Remuneraciones iniciales del resultado(constante)
  sSQL = "INSERT INTO plctsresultado (codcls, pdocts, subcts, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, usrcre, fyhcre) "
  sSQL = sSQL & "SELECT res.codcls, '" & cmbejercicio.Tag & "', '" & cmbPeriodo.Tag & "', res.codpsn, res.codcpc, cxp.secuencia, res.codmon, "
  sSQL = sSQL & "ROUND(ABS(res.importe_mn), 2) AS importe_mn, ROUND(ABS(res.importe_me), 2) AS importe_me, "
  sSQL = sSQL & "codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, res.pdoano, res.pdomes, "
  sSQL = sSQL & "cpc.tipocpc, cpp.impbolecpc, "
  sSQL = sSQL & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
  sSQL = sSQL & "FROM plresultado res "
  sSQL = sSQL & "INNER JOIN plconceplanilla cpp ON res.codcls=cpp.codcls AND res.codcpc=cpp.codcpc "
  sSQL = sSQL & "INNER JOIN plconceproceso cxp ON res.codcls=cxp.codcls AND res.codcpc=cxp.codcpc "
  sSQL = sSQL & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpsn='" & sPersonal & "' "
  sSQL = sSQL & "AND res.pdoano='" & Trim(cmbejercicio) & "' "
  sSQL = sSQL & "AND res.pdomes='" & Left(Trim(cmbPeriodo), 2) & "' "
  sSQL = sSQL & "AND cpp.clasecpc='C' "
  sSQL = sSQL & "AND cxp.codproce='" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "AND IFNULL(cxp.formulafun, '')='' "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "GROUP BY res.codpsn, res.codcpc"
  gdl_Conexion.Execucion sSQL, Inserta

  ' Remuneraciones iniciales del resultado(formulas)
  sSQL = "INSERT INTO plctsresultado (codcls, pdocts, subcts, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, usrcre, fyhcre) "
  sSQL = sSQL & "SELECT res.codcls, '" & cmbejercicio.Tag & "', '" & cmbPeriodo.Tag & "', res.codpsn, res.codcpc, cxp.secuencia, res.codmon, "
  sSQL = sSQL & "ROUND(ABS(res.importe_mn), 2) AS importe_mn, ROUND(ABS(res.importe_me), 2) AS importe_me, "
  sSQL = sSQL & "codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, res.pdoano, res.pdomes, "
  sSQL = sSQL & "cpc.tipocpc, cpp.impbolecpc, "
  sSQL = sSQL & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
  sSQL = sSQL & "FROM plresultado res "
  sSQL = sSQL & "INNER JOIN plconceplanilla cpp ON res.codcls=cpp.codcls AND res.codcpc=cpp.codcpc "
  sSQL = sSQL & "INNER JOIN plconceproceso cxp ON res.codcls=cxp.codcls AND res.codcpc=cxp.codcpc "
  sSQL = sSQL & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSQL = sSQL & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND res.codpsn='" & sPersonal & "'"
  sSQL = sSQL & "AND res.pdoano='" & Trim(cmbejercicio) & "' "
  sSQL = sSQL & "AND res.pdomes='" & Left(Trim(cmbPeriodo), 2) & "' "
  sSQL = sSQL & "AND cpp.clasecpc='F'"
  sSQL = sSQL & "AND cxp.codproce='" & Right(cmbProcesos.Text, 2) & "' "
  sSQL = sSQL & "AND IFNULL(cxp.formulafun, '')='' "
  sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "' "
  sSQL = sSQL & "GROUP BY res.codpsn, res.codcpc"
  gdl_Conexion.Execucion sSQL, Inserta

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
Private Sub ProcesaFormulas(sPersona As String)
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim sValorConcepto As Double
Dim sFormula As String
Dim nLineas As Integer
Dim aLinesFormula() As String
Dim nPosDelmitador As Integer
Dim nVeces As Integer
Dim nI As Integer

' Recuperar Información de imagen de formulas
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
        sFormula = IIf(IsNull(rs!formulafun), "", rs!formulafun)
        nLineas = stNVeces(sFormula, ";")
        nPosDelmitador = 1
  
        If Mid(sFormula, 1, 1) = "@" Then
          ReDim aLinesFormula(1 To nLineas)
          For nVeces = 1 To nLineas
            nPosDelmitador = InStr(1, Trim(sFormula), ";")
            aLinesFormula(nVeces) = Mid(sFormula, 5, nPosDelmitador - 5)
            sFormula = Mid(sFormula, nPosDelmitador + 3)
          Next
          
          For nVeces = 1 To nLineas
            If InStr(1, Trim(aLinesFormula(nVeces)), "@") = 0 Then
              aLinesFormula(nVeces) = CalculaExp(sPersona, rs!codcpc, UCase(aLinesFormula(nVeces)))
              sFormula = aLinesFormula(nVeces)
            Else
              'Reemplazar las Lìneas por Valores Calculados
              For nI = 1 To nLineas
                Do While InStr(1, aLinesFormula(nVeces), "@" + Format(nI, "00"))
                  aLinesFormula(nVeces) = Replace(aLinesFormula(nVeces), "@" + Format(nI, "00"), aLinesFormula(nI), 1, 1)
                Loop
              Next
              aLinesFormula(nVeces) = CalculaExp(sPersona, rs!codcpc, UCase(aLinesFormula(nVeces)))
              sFormula = aLinesFormula(nVeces)
            End If
          Next
        End If
        
        If sFormula <> "" Then
          sValorConcepto = CalculaExp(sPersona, rs!codcpc, UCase(sFormula))
          sSQL = "INSERT INTO plctsresultado "
          sSQL = sSQL & "( "
          sSQL = sSQL & "codcls, "
          sSQL = sSQL & "pdocts, subcts, "
          sSQL = sSQL & "codpsn, "
          sSQL = sSQL & "codcpc, "
          sSQL = sSQL & "secuencia, "
          sSQL = sSQL & "codmon, "
          sSQL = sSQL & "importe_mn, "
          sSQL = sSQL & "importe_me, "
          sSQL = sSQL & "pdoano, pdomes, "
          sSQL = sSQL & "tipocpc, impbolecpc, "
          sSQL = sSQL & "usrcre, fyhcre"
          sSQL = sSQL & ") "
          sSQL = sSQL & "VALUES( "
          sSQL = sSQL & "'" & ps_ClsPlanilla & "', "
          sSQL = sSQL & "'" & cmbejercicio.Tag & "', "
          sSQL = sSQL & "'" & cmbPeriodo.Tag & "', "
          sSQL = sSQL & "'" & sPersona & "', "
          sSQL = sSQL & "'" & rs("codcpc") & "', "
          sSQL = sSQL & "'" & rs("secuencia") & "', "
          sSQL = sSQL & "'" & Choose(sMonedaPago + 1, "N", "E") & "', "
          sSQL = sSQL & "'" & Round(CDec(sValorConcepto), 2) & "', "
          sSQL = sSQL & "'" & Round(CDec(sValorConcepto) / CDec(txtTipoCambio.Text), 2) & "', "
          sSQL = sSQL & "'" & Trim(cmbejercicio.Text) & "', "
          sSQL = sSQL & "'" & Left(Trim(cmbPeriodo.Text), 2) & "', "
          sSQL = sSQL & "'" & rs("tipocpc") & "', "
          sSQL = sSQL & "'" & rs("impbolecpc") & "', "
          sSQL = sSQL & "'" & ps_Usuario & "', "
          sSQL = sSQL & "'" & Format(Now, s_FmtFeHoMysql_0) & "' "
          sSQL = sSQL & ")"
          gdl_Conexion.Execucion sSQL, Inserta
        End If
        rs.MoveNext
    Loop
    rs.Close
End If

End Sub

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
  sSQL = sSQL & "asi.enfermedad AS Dias_Enfermedad, asi.diavacaciones AS Dias_Vacaciones, asi.tercerturno AS Tercer_Turno, asi.diasuspension AS Dias_Suspension, "
  sSQL = sSQL & "asi.licencia AS Licencia, asi.diaferiado AS Dias_Feriado, asi.liquidavacacion AS Vacacion_Pendiente, asi.diagratificacion AS Gratifica_Trunca, asi.opcional AS Valor_Opcional, "
  sSQL = sSQL & "asi.diavacaventa AS Vacacion_Venta, asi.diamediotm AS Dias_MedioTiempo, asi.diaparcial AS Dias_TiempoParcial, asi.horamediotm AS Hora_MedioTiempo, asi.horaparcial AS Hora_TiempoParcial, "
  sSQL = sSQL & "asi.diatradesemanal AS Dias_DescansoSemanal, asi.dialibre AS Dias_Libre, "
  sSQL = sSQL & "cts.numeroanos AS Cts_Anyos, cts.numeromeses AS Cts_Meses, cts.numerodias AS Cts_Dias, " & CDec(txtTipoCambio.Text) & " AS Tipo_Cambio, psn.jornadalaboral AS Jornada_laboral, "
  sSQL = sSQL & "'N' AS CtaCte_Gratifica, 'N' AS CtaCte_Descuento, "
  sSQL = sSQL & "IF(psn.chk27252='" & s_Estado_Act & "', 'S', 'N') AS L27252, "
  sSQL = sSQL & "IF(asi.indvacadelanta='" & s_Estado_Act & "', 'S', 'N') AS Vacacion_Adelantada, "
  sSQL = sSQL & "asi.diavacavencida AS Vacacion_Vencida, "
  sSQL = sSQL & "(YEAR('" & Format(mskFecha, "yyyy-mm-dd") & "')-YEAR(psn.fecnacimiento) + IF(DATE_FORMAT('" & Format(mskFecha, s_FormatoFecha) & "','%m-%d') > DATE_FORMAT(psn.fecnacimiento,'%m-%d'),0,-1)) AS Edad_Personal "
  sSQL = sSQL & "FROM plpersonal psn "
  sSQL = sSQL & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
  sSQL = sSQL & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
  sSQL = sSQL & "LEFT JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn AND asi.codpdo='" & sPeriodo & "' "
  sSQL = sSQL & "LEFT JOIN plctsmovimiento cts ON psn.codcls=cts.codcls AND psn.codpsn=cts.codpsn AND cts.pdocts='" & cmbejercicio.Tag & "' AND cts.subcts='" & cmbPeriodo.Tag & "' "
  sSQL = sSQL & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  sSQL = sSQL & "AND psn.codpsn='" & sPersona & "'"
  Set porstVariables = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, sSQL)

End Sub
Private Function FAcumulado(ByVal sConcepto As String, ByVal sDelMes As String) As Double
  FAcumulado = oCalculo.Acumulado(Mid(sConcepto, 2), sDelMes)
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

Private Function FSi(sCondicion1, sOperador, sCondicion2, sTipoComparacion, sVerdadero, sFalso) As Double

'Comparaciòn Caracteres
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

'Comparaciòn Números
If sTipoComparacion = "N" Then
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

'Comparación Fechas
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

'Comparaciòn Lógico
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

End Function
Private Function FUltimoMovimiento(ByVal sConcepto As String, ByVal sTblResultado As String, ByVal nMesAnterior As Integer, Optional ByVal sTipoProceso As String) As Double
  sTblResultado = "pl" & IIf(sTblResultado = "C", "cts", "") & "resultado"
  FUltimoMovimiento = oCalculo.UltimoMovimiento(Mid(sConcepto, 2), sTblResultado, nMesAnterior, sTipoProceso)
End Function
Private Function FVacaTruncas() As Double
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
  If Right(mskFecha.ClipText, 4) <> Trim(cmbejercicio) Then Beep: MsgBox "Fecha debe ser del periodo a Procesar", vbExclamation: mskFecha.SetFocus: Cancel = True: Exit Sub
  If Mid(mskFecha.ClipText, 3, 2) <> Left(Trim(cmbPeriodo), 2) Then Beep: MsgBox "Fecha debe ser del mes a Procesar", vbExclamation: mskFecha.SetFocus: Cancel = True: Exit Sub
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
