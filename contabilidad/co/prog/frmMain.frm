VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   3885
   ClientLeft      =   2505
   ClientTop       =   1455
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   8715
   Visible         =   0   'False
   Begin Crystal.CrystalReport rptMain 
      Left            =   1575
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox uctEstado1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   8655
      TabIndex        =   8
      Top             =   3585
      Width           =   8715
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   8655
      TabIndex        =   1
      Top             =   2010
      Width           =   8715
      Begin VB.Label LblTitu 
         Caption         =   "Usuario: "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblVar 
         Caption         =   "<usuario>"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LblTitu 
         Caption         =   "Mes:"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label lblVar 
         Caption         =   "<Empresa> 1234567890 123456789 123456789 123456789 1234567890"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   9855
      End
      Begin VB.Label lblVar 
         Caption         =   "<Mes>"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblVar 
         Caption         =   "<sistema>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblVar 
         Caption         =   "<Año>"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label LblTitu 
         Caption         =   "Período:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Diario"
            Object.ToolTipText     =   "Diario"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mayor"
            Object.ToolTipText     =   "Mayor Auxiliar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Balance"
            Object.ToolTipText     =   "Balance de Comprobación"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Auxiliar"
            Object.ToolTipText     =   "Auxiliares"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A01A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A12E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlDialogo 
      Left            =   2400
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "&Transacciones"
      Begin VB.Menu opcTCpr 
         Caption         =   "&Compras"
      End
      Begin VB.Menu opcTVta 
         Caption         =   "&Ventas"
      End
      Begin VB.Menu opcTHPr 
         Caption         =   "&Honorarios"
      End
      Begin VB.Menu divT1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu opcTBan 
         Caption         =   "Caja  &Bancos"
      End
      Begin VB.Menu divT2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu opcTCpb 
         Caption         =   "&Diario"
      End
      Begin VB.Menu opcRCpbNCu 
         Caption         =   "Comprobantes &No Cuadrados"
      End
      Begin VB.Menu divT3 
         Caption         =   "-"
      End
      Begin VB.Menu opcTPdo 
         Caption         =   "&Pedido de Compra"
      End
      Begin VB.Menu opcTCon 
         Caption         =   "Contrato de &Servicios"
      End
      Begin VB.Menu divT4 
         Caption         =   "-"
      End
      Begin VB.Menu opcTRteVta 
         Caption         =   "Ventas &Recurrentes"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuDro 
         Caption         =   "&Diarios"
         Begin VB.Menu opcRDroAux 
            Caption         =   "Diario &Auxiliar"
         End
         Begin VB.Menu opcRDroGrl 
            Caption         =   "Diario &General"
         End
      End
      Begin VB.Menu mnuRMay 
         Caption         =   "Mayores"
         Begin VB.Menu opcRMayAux 
            Caption         =   "Mayor &Auxiliar"
         End
         Begin VB.Menu opcRMayGrl 
            Caption         =   "Mayor &General"
         End
      End
      Begin VB.Menu mnuRCja 
         Caption         =   "Caja & Bancos"
         Begin VB.Menu opcRLbrCja 
            Caption         =   "&Libro Caja Bancos"
         End
         Begin VB.Menu opcRMovCja 
            Caption         =   "&Movimiento Caja Bancos"
         End
         Begin VB.Menu divR0 
            Caption         =   "-"
         End
         Begin VB.Menu opcRMovBco 
            Caption         =   "Movimiento &Bancos"
         End
      End
      Begin VB.Menu opcRBceCpb 
         Caption         =   "&Balance Comprobación"
      End
      Begin VB.Menu opcREFi 
         Caption         =   "Estados Financieros"
      End
      Begin VB.Menu divR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCCo 
         Caption         =   "Centros de Costos"
         Begin VB.Menu opcRMayAuxCCo 
            Caption         =   "&Mayor Aux. por C.Costo"
         End
         Begin VB.Menu opcRBceCpbCCo 
            Caption         =   "Balance de Comprobación por C.Costo"
         End
         Begin VB.Menu opcREFiCCo 
            Caption         =   "Ganancias y Pérdidas por C.Costo"
         End
         Begin VB.Menu opcREFiCCoMes 
            Caption         =   "Ganancias y Pérdidas Mes a Mes por C.Costo"
         End
         Begin VB.Menu opcRGtoCCo 
            Caption         =   "Gastos Mes a Mes por C.Costo"
         End
      End
      Begin VB.Menu divR2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRSdo 
         Caption         =   "&Saldos"
         Begin VB.Menu opcRSdoMes 
            Caption         =   "Saldos &Mes a Mes"
         End
         Begin VB.Menu opcRSdoMesCta 
            Caption         =   "Saldos Mes a Mes por &Cuenta"
         End
         Begin VB.Menu opcRSdoMesAux 
            Caption         =   "Saldos Mes a Mes por &Auxiliar"
         End
      End
      Begin VB.Menu divR3 
         Caption         =   "-"
      End
      Begin VB.Menu opcRFluCjaRea 
         Caption         =   "Flujo de Caja Real"
      End
      Begin VB.Menu opcRFluEfectivo 
         Caption         =   "Estado de Flujo de Efectivo"
      End
      Begin VB.Menu opcRCtlPsp 
         Caption         =   "Control de Presupuestos"
      End
      Begin VB.Menu opcRCerIR4 
         Caption         =   "Certificado de 4ª Categoría"
      End
      Begin VB.Menu divR4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRReg 
         Caption         =   "Reg&istros"
         Begin VB.Menu opcRRegCpr 
            Caption         =   "Registro de &Compras"
         End
         Begin VB.Menu opcRRegVta 
            Caption         =   "Registro de &Ventas"
         End
         Begin VB.Menu opcRRegHPr 
            Caption         =   "Registro de &Honorarios"
         End
         Begin VB.Menu divR5 
            Caption         =   "-"
         End
         Begin VB.Menu opcRRegRtc 
            Caption         =   "Registro de Retenciones"
         End
         Begin VB.Menu opcRRegPcp 
            Caption         =   "Registro de Percepciones"
         End
      End
      Begin VB.Menu divR6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCtaCte 
         Caption         =   "C&uentas Corrientes"
         Begin VB.Menu opcRCCtAux 
            Caption         =   "Ctas.Ctes. por &Auxiliar"
         End
         Begin VB.Menu opcRCCtCta 
            Caption         =   "Ctas.Ctes. por Cuenta"
         End
         Begin VB.Menu opcRCCtPHs 
            Caption         =   "Ctas.Ctes. por Cuenta Detalle"
         End
         Begin VB.Menu opcRCCtHst 
            Caption         =   "Ctas.Ctes. &Histórico"
         End
         Begin VB.Menu opcRCCtAtgSdo 
            Caption         =   "Ctas.Ctes. por Antiguedad de &Saldos"
         End
         Begin VB.Menu opcRCCtCCo 
            Caption         =   "Ctas.Ctes. por Centro de C&ostos"
         End
      End
      Begin VB.Menu divR7 
         Caption         =   "-"
      End
      Begin VB.Menu opcRCCtPdo 
         Caption         =   "Análisis de Pedidos"
      End
      Begin VB.Menu divR8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRTipo54 
         Caption         =   "Reportes &Tipo 054"
         Begin VB.Menu opcRTp54Cpr 
            Caption         =   "Registro de Compras"
         End
         Begin VB.Menu opcRTp54Vta 
            Caption         =   "Registro de Ventas"
         End
         Begin VB.Menu divR9 
            Caption         =   "-"
         End
         Begin VB.Menu opcRTp54Ret 
            Caption         =   "Reg. Retención y Percepción"
         End
      End
      Begin VB.Menu mnuRTipo56 
         Caption         =   "Reportes &Tipo 056"
         Begin VB.Menu opcRTp56Cab 
            Caption         =   "&Cabecera Reporte"
         End
         Begin VB.Menu divR10 
            Caption         =   "-"
         End
         Begin VB.Menu opcRTp56Cpr 
            Caption         =   "Registro de &Compras"
         End
         Begin VB.Menu opcRTp56Vta 
            Caption         =   "Regsitro de Ventas"
         End
         Begin VB.Menu opcRTp56Hpr 
            Caption         =   "Registro de Honorarios Profesionales"
         End
         Begin VB.Menu divR11 
            Caption         =   "-"
         End
         Begin VB.Menu opcRTp56Dro 
            Caption         =   "Diario Auxiliar"
         End
         Begin VB.Menu opcRTp56May 
            Caption         =   "Mayor Auxiliar"
         End
         Begin VB.Menu opcRTp56Bce 
            Caption         =   "Balance General"
         End
         Begin VB.Menu opcRTp56EFi 
            Caption         =   "Estados Financieros"
         End
         Begin VB.Menu divR12 
            Caption         =   "-"
         End
         Begin VB.Menu opcRTp56CIR 
            Caption         =   "Certificado de Renta"
         End
         Begin VB.Menu divR13 
            Caption         =   "-"
         End
         Begin VB.Menu opcrpt56CV 
            Caption         =   "Compras y Ventas Txt"
         End
      End
      Begin VB.Menu opcLibros 
         Caption         =   "Libros Oficiales 2009"
      End
      Begin VB.Menu divR14 
         Caption         =   "-"
      End
      Begin VB.Menu opcRCCtDetra 
         Caption         =   "Análisis de Detracciones"
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "&Procesos"
      Begin VB.Menu opcPDifTCb 
         Caption         =   "&Diferencia por Tipo de Cambio"
      End
      Begin VB.Menu opcPMay 
         Caption         =   "&Mayorización"
      End
      Begin VB.Menu opcPeli 
         Caption         =   "&Eliminar Asientos Automaticos"
      End
      Begin VB.Menu divP1 
         Caption         =   "-"
      End
      Begin VB.Menu opcPTrfPDT 
         Caption         =   "&Transferencia al PDT"
         Begin VB.Menu opcPPDTHPrMes 
            Caption         =   "&Honorarios Profesionales"
         End
         Begin VB.Menu opcPPDTDAOT 
            Caption         =   "&DAOT"
         End
         Begin VB.Menu opcPPDTRet 
            Caption         =   "Registro de Retenciones - Percepciones"
         End
         Begin VB.Menu opcPPDTEeFf 
            Caption         =   "Estados Financieros / Balance Comprobación"
         End
         Begin VB.Menu opcPPDTVta 
            Caption         =   "&Ventas Anuales"
         End
         Begin VB.Menu opcPPDB 
            Caption         =   "Informacion PDB"
         End
         Begin VB.Menu opcPPDTDetra 
            Caption         =   "Cuenta Detracción"
         End
      End
      Begin VB.Menu divP2 
         Caption         =   "-"
      End
      Begin VB.Menu opcPcieApe 
         Caption         =   "Asientos de &Cierre/Apertura"
      End
      Begin VB.Menu divP3 
         Caption         =   "-"
      End
      Begin VB.Menu opcPFil 
         Caption         =   "&Archivos de Información"
      End
   End
   Begin VB.Menu mnuMaestros 
      Caption         =   "Ta&blas"
      Begin VB.Menu opcMDro 
         Caption         =   "&Diarios"
      End
      Begin VB.Menu opcMCta 
         Caption         =   "&Plan de Cuentas"
      End
      Begin VB.Menu opcMCCo 
         Caption         =   "&Centros de Costos"
      End
      Begin VB.Menu opcMEfe 
         Caption         =   "Flujo de &Efectivo"
      End
      Begin VB.Menu opcMFjo 
         Caption         =   "&Flujo de Caja"
      End
      Begin VB.Menu opcMBco 
         Caption         =   "&Bancos"
      End
      Begin VB.Menu opcMPago 
         Caption         =   "Medios de Pago"
      End
      Begin VB.Menu divM1 
         Caption         =   "-"
      End
      Begin VB.Menu opcMAux 
         Caption         =   "&Auxiliares"
      End
      Begin VB.Menu opcMTDc 
         Caption         =   "&Tipos de Documentos"
      End
      Begin VB.Menu opcMTCb 
         Caption         =   "Tipos de Cam&bio"
      End
      Begin VB.Menu opcMTCbCie 
         Caption         =   "Tipos de Cambio de Ci&erre"
      End
      Begin VB.Menu divM2 
         Caption         =   "-"
      End
      Begin VB.Menu opcMPsp 
         Caption         =   "Pres&upuestos"
      End
      Begin VB.Menu divM3 
         Caption         =   "-"
      End
      Begin VB.Menu opcMEFi 
         Caption         =   "&Estados Financieros"
      End
      Begin VB.Menu opcMFil 
         Caption         =   "&Plantilla de Archivos"
      End
      Begin VB.Menu divM4 
         Caption         =   "-"
      End
      Begin VB.Menu opcMAsi 
         Caption         =   "Asiento &Tipo"
      End
      Begin VB.Menu opcMDpe 
         Caption         =   "Pro&yectos"
      End
      Begin VB.Menu OpcMDupli 
         Caption         =   "Asi&ento Tipo-Diario"
      End
      Begin VB.Menu OpcMDetra 
         Caption         =   "Detraccion"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "&Utilitarios"
      Begin VB.Menu opcOEmpAct 
         Caption         =   "&Empresa/Año Activos"
      End
      Begin VB.Menu opcOMesAct 
         Caption         =   "&Mes Activo"
      End
      Begin VB.Menu divU1 
         Caption         =   "-"
      End
      Begin VB.Menu opcMEmp 
         Caption         =   "&Nueva Empresa/Año"
      End
      Begin VB.Menu opcOCieMes 
         Caption         =   "&Cierre de Meses"
      End
      Begin VB.Menu divU2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilCfg 
         Caption         =   "Configuración"
         Begin VB.Menu opcOMesAtu 
            Caption         =   "&Mes Actual"
         End
         Begin VB.Menu opcOPar 
            Caption         =   "&Parámetros"
         End
      End
      Begin VB.Menu divU3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilSeg 
         Caption         =   "&Seguridad"
         Begin VB.Menu opcMUsr 
            Caption         =   "&Usuarios"
         End
         Begin VB.Menu opcMMdl 
            Caption         =   "&Módulos (((DESARROLLO)))"
            Visible         =   0   'False
         End
         Begin VB.Menu opcMPms 
            Caption         =   "&Permisos"
         End
         Begin VB.Menu opcRBitacor 
            Caption         =   "Administración Bitacora"
         End
      End
      Begin VB.Menu opcOCla 
         Caption         =   "Cambio de Cla&ve"
      End
      Begin VB.Menu divU4 
         Caption         =   "-"
      End
      Begin VB.Menu opcPBackup 
         Caption         =   "&Backup de Información"
      End
      Begin VB.Menu opcPTraInf 
         Caption         =   "&Transferencia de Información"
      End
      Begin VB.Menu divU5 
         Caption         =   "-"
      End
      Begin VB.Menu opcSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu opcAcerca 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Buscar Ayuda acerca de..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[REVISAR: ¿Necesario?
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
']REVISAR.

Private pocnnBDC As ADODB.Connection
Private porstPermisos As ADODB.Recordset

Private Sub Form_Load()

   Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
   Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
   Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
   Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
   
   lblVar(0) = gsRazEmp
   lblVar(1) = gsNomSis
   lblVar(2) = gsAnoAct
   lblVar(3) = gsMesAct
   lblVar(4) = gsCodUsr '2015-09-21 mod frm menu usu y rsocial

   Me.opcMMdl.Visible = IIf(gbEsUsr, False, True)

   Set pocnnBDC = New Connection
   Set porstPermisos = New Recordset
   pocnnBDC.CursorLocation = adUseClient
   pocnnBDC.ConnectionString = CONNSTRG & gsNomBDC
   porstPermisos.Source = "SELECT CodMdl, IndPms01, IndPms02, IndPms03, IndPms04, IndPms05, IndPms06, "
   porstPermisos.Source = porstPermisos.Source & "IndPms07, IndPms08, IndPms09, IndPms10 "
   porstPermisos.Source = porstPermisos.Source & "FROM SGPms "
   porstPermisos.Source = porstPermisos.Source & "WHERE CodUsr='" & gsCodUsr & "' AND CodEmp='" & gsCodEmp & "' AND CodSis='" & gsCodSis & "'"
   porstPermisos.CursorType = adOpenStatic
   porstPermisos.LockType = adLockReadOnly

End Sub

Private Sub Form_Activate()
'ini 2015-05-18 validacion frm
'ahora que hice la correccion en proc/dif cambio
'para que salga del load, salio un error
'pongo esto para subsanarlo
'fRstClose porstPermisos
'fCnnClose pocnnBDC
    If porstPermisos.State = adStateOpen Then porstPermisos.Close
    If pocnnBDC.State = adStateOpen Then pocnnBDC.Close
'fin 2015-05-18 validacion frm

  pocnnBDC.ConnectionString = CONNSTRG & gsNomBDC
  pocnnBDC.Open
  porstPermisos.ActiveConnection = pocnnBDC
  porstPermisos.Open
End Sub

Private Sub Form_Deactivate()
   porstPermisos.Close
   pocnnBDC.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim dnContador As Integer

   For dnContador = Forms.Count - 1 To 1 Step -1
      Unload Forms(dnContador)
   Next
   
'[REVISAR: ¿Necesario?
   If Me.WindowState <> vbMinimized Then
      SaveSetting App.Title, "Settings", "MainLeft", Me.Left
      SaveSetting App.Title, "Settings", "MainTop", Me.Top
      SaveSetting App.Title, "Settings", "MainWidth", Me.Width
      SaveSetting App.Title, "Settings", "MainHeight", Me.Height
   End If
']REVISAR.

End Sub

Private Sub frmMCpbGrd_Click()

End Sub

Private Sub opcAcerca_Click()
fAcercaDe.Show
End Sub

Private Sub opcLibros_Click()
    ppCapturaTitulo frmLibros, opcLibros
End Sub

Private Sub opcMAsi_Click()
  Call ppCapturaTitulo(frmMAsiGrd, opcMAsi)
End Sub

Private Sub opcMBco_Click()
  ppCapturaTitulo frmMBcoGrd, opcMBco
End Sub

Private Sub OpcMDetra_Click()
'frmMDetracGrd.Show
  ppCapturaTitulo frmMDetracGrd, OpcMDetra

End Sub

Private Sub opcMDpe_Click()
  ppCapturaTitulo frmMDPeGrd, opcMDpe
End Sub

Private Sub OpcMDupli_Click()
Call ppCapturaTitulo(frmMCpbGrd, OpcMDupli)
End Sub

Private Sub opcMEfe_Click()
  ppCapturaTitulo frmMEfeGrd, opcMEfe
End Sub
Private Sub opcMFil_Click()
    ppCapturaTitulo frmMFilGrd, opcMFil
End Sub
Private Sub opcMFjo_Click()
   Call ppCapturaTitulo(frmMFjoGrd, opcMFjo)
End Sub

Private Sub opcMPago_Click()
    ppCapturaTitulo frmMpago, opcMPago
End Sub

Private Sub opcPBackup_Click()
   Call ppCapturaTitulo(frmPBackup, opcPBackup)
End Sub

Private Sub opcPeli_Click()
   Call ppCapturaTitulo(frmPeli, opcPeli)
End Sub

Private Sub opcPFil_Click()
  ppCapturaTitulo frmPFil, opcPFil
End Sub

Private Sub opcPPDB_Click()
  ppCapturaTitulo frmPPDB, opcPPDB
End Sub

Private Sub opcPPDTDetra_Click()
  ppCapturaTitulo frmPPDTDetra, opcPPDTDetra
End Sub

Private Sub opcPPDTEeFf_Click()
   Call ppCapturaTitulo(frmPPDTEeFf, opcPPDTEeFf)
End Sub
Private Sub opcPPDTVta_Click()
  ppCapturaTitulo frmPPDTVta, opcPPDTVta
End Sub
Private Sub opcPTraInf_Click()
    Call ppCapturaTitulo(frmPTraInf, opcPTraInf)
End Sub
Private Sub opcRBitacor_Click()
    Call ppCapturaTitulo(frmRBitacor, opcRBitacor)
End Sub

Private Sub opcRCCtDetra_Click()
  ppCapturaTitulo frmRCCtDetra, opcRCCtDetra
End Sub

Private Sub opcRCCtPdo_Click()
  Call ppCapturaTitulo(frmRCCtPdo, opcRCCtPdo)
End Sub

Private Sub opcRCCtPHs_Click()
  Call ppCapturaTitulo(frmRCCtPHs, opcRCCtPHs)
End Sub

Private Sub opcRFluEfectivo_Click()
    Call ppCapturaTitulo(frmRFluEfectivo, opcRFluEfectivo)
End Sub

Private Sub opcRGtoCCo_Click()
   ppCapturaTitulo frmRGtoCCo, opcRGtoCCo
End Sub

Private Sub opcRMovBco_Click()
   ppCapturaTitulo frmRMovBco, opcRMovBco
End Sub

Private Sub opcRMovCja_Click()
   Call ppCapturaTitulo(frmRMovCja, opcRMovCja)
End Sub


Private Sub opcRTp54Cpr_Click()
  ppCapturaTitulo frmRTp54Cpr, opcRTp54Cpr
End Sub

Private Sub opcRTp54Ret_Click()
  ppCapturaTitulo frmRTp54Ret, opcRTp54Ret
End Sub

Private Sub opcRTp54Vta_Click()
  ppCapturaTitulo frmRTp54Vta, opcRTp54Vta
End Sub

Private Sub opcRTp56Bce_Click()
  ppCapturaTitulo frmRTp56Bce, opcRTp56Bce
End Sub

Private Sub opcRTp56Cab_Click()
  ppCapturaTitulo frmRTp56Cab, opcRTp56Cab
End Sub

Private Sub opcRTp56CIR_Click()
  ppCapturaTitulo frmRTp56CIR, opcRTp56CIR
End Sub

Private Sub opcRTp56Cpr_Click()
  ppCapturaTitulo frmRTp56Cpr, opcRTp56Cpr
End Sub

Private Sub opcrpt56CV_Click()
  ppCapturaTitulo frmrpt56CV, opcrpt56CV
End Sub

Private Sub opcRTp56Dro_Click()
  ppCapturaTitulo frmRTp56Dro, opcRTp56Dro
End Sub

Private Sub opcRTp56EFi_Click()
  ppCapturaTitulo frmRTp56EFi, opcRTp56EFi
End Sub

Private Sub opcRTp56HPr_Click()
  ppCapturaTitulo frmRTp56HPr, opcRTp56Hpr
End Sub

Private Sub opcRTp56May_Click()
  ppCapturaTitulo frmRTp56May, opcRTp56May
End Sub

Private Sub opcRTp56Vta_Click()
  ppCapturaTitulo frmRTp56Vta, opcRTp56Vta
End Sub

Private Sub opcTBan_Click()
   If gsMesAct <> "00" And gsMesAct <> "13" Then ppCapturaTitulo frmTBanGrd, opcTBan
End Sub

Private Sub opcTCon_Click()
  If gsMesAct <> "00" And gsMesAct <> "13" Then Call ppCapturaTitulo(frmTConGrd, opcTCon)
End Sub

Private Sub opcTCpr_Click()
   If gsMesAct <> "00" And gsMesAct <> "13" Then Call ppCapturaTitulo(frmTCprGrd, opcTCpr)
End Sub

Private Sub opcTPdo_Click()
  Call ppCapturaTitulo(frmTPdoGrd, opcTPdo)
End Sub

Private Sub opcTVta_Click()
   If gsMesAct <> "00" And gsMesAct <> "13" Then Call ppCapturaTitulo(frmTVtaGrd, opcTVta)
End Sub

Private Sub opcTHPr_Click()
   If gsMesAct <> "00" And gsMesAct <> "13" Then Call ppCapturaTitulo(frmTHPrGrd, opcTHPr)
End Sub

Private Sub opcTCpb_Click()
  ppCapturaTitulo frmTCpbGrd, opcTCpb
End Sub

Private Sub opcTRteVta_Click()
  If gsMesAct <> "00" And gsMesAct <> "13" Then Call ppCapturaTitulo(frmTRteVtaGrd, opcTRteVta)
End Sub

Private Sub opcRCpbNCu_Click()
   Call ppCapturaTitulo(frmRCpbNCu_1, opcRCpbNCu)
End Sub

Private Sub opcRBceCpb_Click()
   Call ppCapturaTitulo(frmRBceCpb, opcRBceCpb)
End Sub

Private Sub opcRCCtAux_Click()
   Call ppCapturaTitulo(frmRCCtAux, opcRCCtAux)
End Sub

Private Sub opcRCCtCta_Click()
   Call ppCapturaTitulo(frmRCCtCta, opcRCCtCta)
End Sub

Private Sub opcRBceCpbCCo_Click()
   Call ppCapturaTitulo(frmRBceCpbCCo, opcRBceCpbCCo)
End Sub

Private Sub opcRCCtAtgSdo_Click()
   Call ppCapturaTitulo(frmRCCtAtgSdo, opcRCCtAtgSdo)
End Sub

Private Sub opcRCCtCCo_Click()
   Call ppCapturaTitulo(frmRCCtCCo, opcRCCtCCo)
End Sub

Private Sub opcRCerIR4_Click()
   Call ppCapturaTitulo(frmRCerIR4, opcRCerIR4)
End Sub

Private Sub opcRCNS_Click()
'   Call ppCapturaTitulo(frmRCNS, opcRCNS)
End Sub

Private Sub opcRCCtHst_Click()
   Call ppCapturaTitulo(frmRCCtHst, opcRCCtHst)
End Sub

Private Sub opcRCtlPsp_Click()
   Call ppCapturaTitulo(frmRCtlPsp, opcRCtlPsp)
End Sub

Private Sub opcRDroAux_Click()
   Call ppCapturaTitulo(frmRDroAux, opcRDroAux)
End Sub

Private Sub opcRDroGrl_Click()
   Call ppCapturaTitulo(frmRDroGrl, opcRDroGrl)
End Sub

Private Sub opcREFi_Click()
   Call ppCapturaTitulo(frmREFi, opcREFi)
End Sub

Private Sub opcREFiCCo_Click()
   Call ppCapturaTitulo(frmREFiCCo, opcREFiCCo)
End Sub

Private Sub opcREFiCCoMes_Click()
   Call ppCapturaTitulo(frmREFiCCoMes, opcREFiCCoMes)
End Sub

Private Sub opcRFluCjaRea_Click()
   Call ppCapturaTitulo(frmRFluCjaRea, opcRFluCjaRea)
End Sub

Private Sub opcRLbrCja_Click()
   Call ppCapturaTitulo(frmRLbrCja, opcRLbrCja)
End Sub

Private Sub opcRMayAux_Click()
   Call ppCapturaTitulo(frmRMayAux, opcRMayAux)
End Sub

Private Sub opcRMayAuxCCo_Click()
   Call ppCapturaTitulo(frmRMayAuxCCo, opcRMayAuxCCo)
End Sub

Private Sub opcRMayGrl_Click()
   Call ppCapturaTitulo(frmRMayGrl, opcRMayGrl)
End Sub

Private Sub opcRRegCpr_Click()
   Call ppCapturaTitulo(frmRRegCpr, opcRRegCpr)
End Sub

Private Sub opcRRegHPr_Click()
   Call ppCapturaTitulo(frmRRegHPr, opcRRegHPr)
End Sub

Private Sub opcRRegPcp_Click()
   Call ppCapturaTitulo(frmRRegPcp, opcRRegPcp)
End Sub

Private Sub opcRRegRtc_Click()
   Call ppCapturaTitulo(frmRRegRtc, opcRRegRtc)
End Sub

Private Sub opcRRegVta_Click()
   Call ppCapturaTitulo(frmRRegVta, opcRRegVta)
End Sub

Private Sub opcRSdoMes_Click()
   Call ppCapturaTitulo(frmRSdoMes, opcRSdoMes)
End Sub

Private Sub opcRSdoMesAux_Click()
   Call ppCapturaTitulo(frmRSdoMesAux, opcRSdoMesAux)
End Sub

Private Sub opcRSdoMesCta_Click()
   Call ppCapturaTitulo(frmRSdoMesCta, opcRSdoMesCta)
End Sub

Private Sub opcPMay_Click()
   Call ppCapturaTitulo(frmPMay, opcPMay)
End Sub

Private Sub opcPDifTCb_Click()
  Call ppCapturaTitulo(frmPDifTCb, opcPDifTCb)
End Sub

Private Sub opcPPDTHPrMes_Click()
   Call ppCapturaTitulo(frmPPDTHPrMes, opcPPDTHPrMes)
End Sub

Private Sub opcPPDTDAOT_Click()
   Call ppCapturaTitulo(frmPPDTDAOT, opcPPDTDAOT)
End Sub

Private Sub opcPPDTRet_Click()
   Call ppCapturaTitulo(frmPPDTRet, opcPPDTRet)
End Sub

Private Sub opcPcieApe_Click()
   Call ppCapturaTitulo(frmPCieApe, opcPcieApe)
End Sub
Private Sub opcMAux_Click()
   Call ppCapturaTitulo(frmMAuxGrd, opcMAux)
End Sub

Private Sub opcMCCo_Click()
   Call ppCapturaTitulo(frmMCCoGrd, opcMCCo)
End Sub

Private Sub opcMCta_Click()
   Call ppCapturaTitulo(frmMCtaGrd, opcMCta)
End Sub

Private Sub opcMDro_Click()
   Call ppCapturaTitulo(frmMDroGrd, opcMDro)
End Sub

Private Sub opcMTCb_Click()
   Call ppCapturaTitulo(frmMTCbGrd, opcMTCb)
End Sub

Private Sub opcMTCbCie_Click()
   Call ppCapturaTitulo(frmMTCbCie, opcMTCbCie)
End Sub

Private Sub opcMTDc_Click()
   Call ppCapturaTitulo(frmMTDcGrd, opcMTDc)
End Sub

Private Sub opcMEFi_Click()
   Call ppCapturaTitulo(frmMEFiGrd, opcMEFi)
End Sub

Private Sub opcMPsp_Click()
   Call ppCapturaTitulo(frmMPspGrd, opcMPsp)
End Sub

Private Sub opcOCieMes_Click()
   Call ppCapturaTitulo(frmOCieMes, opcOCieMes)
End Sub

Private Sub opcOEmpAct_Click()
   Call ppCapturaTitulo(frmOEmpAct, opcOEmpAct)
End Sub

Private Sub opcOMesAct_Click()
   Call ppCapturaTitulo(frmOMesAct, opcOMesAct)
End Sub

Private Sub opcOMesAtu_Click()
   Call ppCapturaTitulo(frmOMesAtu, opcOMesAtu)
End Sub

Private Sub opcOPar_Click()
   Call ppCapturaTitulo(frmOPar, opcOPar)
End Sub

Private Sub opcMEmp_Click()
   Call ppCapturaTitulo(frmMEmpGrd, opcMEmp)
End Sub

Private Sub opcMUsr_Click()
   Call ppCapturaTitulo(frmMUsrGrd, opcMUsr)
End Sub

Private Sub opcMMdl_Click()
   Call ppCapturaTitulo(frmMMdlGrd, opcMMdl)
End Sub

Private Sub opcMPms_Click()
  ' Call ppCapturaTitulo(frmMPmsGrd, opcMPms)
  ' modificado nuevo menu de accesos a usuarios
   Call ppCapturaTitulo(frmMSeg, opcMPms)
End Sub

Private Sub opcOCla_Click()
   Call ppCapturaTitulo(frmOCla, opcOCla)
End Sub

Private Sub opcSalir_Click()
   Unload Me
End Sub

'[REVISAR: Definir si se usa.
'Private Sub mnuHelpAbout_Click()
'   frmAbout.Show vbModal, Me
'End Sub

'Private Sub mnuHelpSearchForHelpOn_Click()
'   Dim nRet As Integer
   
   'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
   'puede establecer el archivo de Ayuda para su aplicación en el cuadro
   'de diálogo Propiedades del proyecto
'   If Len(App.HelpFile) = 0 Then
'      MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
'   Else
'      On Error Resume Next
'      nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'      If Err Then
'         MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
'      End If
'   End If

'End Sub

'Private Sub mnuHelpContents_Click()
'   Dim nRet As Integer

   'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
   'puede establecer el archivo de Ayuda para la aplicación en el cuadro
   'de diálogo Propiedades del proyecto
'   If Len(App.HelpFile) = 0 Then
'      MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
'   Else
'      On Error Resume Next
'      nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'      If Err Then
'         MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
'      End If
'   End If

'End Sub
']REVISAR.

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
   On Error Resume Next
   
'[ARREGLAR: Difinir procedimiento.
   Select Case Button.Key
      Case "Diario"
'         opcTRegDro_Click
      Case "Mayor"
'         opc...
      Case "Balance"
'         opc...
      Case "Auxiliar"
'         opcMAux_Click
   End Select
']ARREGLAR.
End Sub

Sub ppCapturaTitulo(tofrm1 As Form, tomnu1 As Menu)
'   On Error GoTo Err

   pExitForm = 0 '2015-05-18 validacion
   
   Dim dnPos As Integer, dsCad As String
   
   If gbEsUsr Then
      With porstPermisos
         If .RecordCount = 0 Then
            MsgBox Choose(gsIdioma, "Opción restringida", "Restricted Option"), vbInformation
            Exit Sub
         End If
         .MoveFirst
         .Find "CodMdl='" & tofrm1.Name & "'"
         If .EOF Then
            MsgBox Choose(gsIdioma, "Opción restringida", "Restricted Option"), vbInformation
            Exit Sub
         End If
         gbPms01 = !IndPms01
         gbPms02 = !IndPms02
         gbPms03 = !IndPms03
         gbPms04 = !IndPms04
         gbPms05 = !IndPms05
         gbPms06 = !IndPms06
         gbPms07 = !IndPms07
         gbPms08 = !IndPms08
         gbPms09 = !IndPms09
         gbPms10 = !IndPms10
      End With
   Else
      gbPms01 = True
      gbPms02 = True
      gbPms03 = True
      gbPms04 = True
      gbPms05 = True
      gbPms06 = True
      gbPms07 = True
      gbPms08 = True
      gbPms09 = True
      gbPms10 = True
   End If
   
   With tofrm1
      dsCad = tomnu1.Caption
      dnPos = InStr(dsCad, "&")
      If Not dnPos = 0 Then
         dsCad = Mid(dsCad, 1, dnPos - 1) & Mid(dsCad, dnPos + 1)
      End If
      
      Select Case Mid(tomnu1.Name, 4, 1)
      Case "L"
         .Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & dsCad
      Case "M"
         .Caption = dsCad
      Case "P"
         .Caption = Choose(gsIdioma, "Proceso de ", "Process of ") & dsCad
      Case "E"
         .Caption = dsCad
      Case "R"
'         .Caption = "Reporte de " & dsCad
         .Caption = dsCad
      Case "T"
         .Caption = dsCad & " (" & gsMesAct & "/" & gsAnoAct & ")"
      Case "O"
         .Caption = dsCad
      End Select
'ini 2015-05-18 validacion frm
'      .Show 'vbModal
      If pExitForm = 0 Then
        .Show 'vbModal
      End If
      If pExitForm = 1 Then
      Unload tofrm1
         'vbModal
      End If
'fin 2015-05-18 validacion frm
   End With

   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

