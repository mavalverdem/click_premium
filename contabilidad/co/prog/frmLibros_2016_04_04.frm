VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLibros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros y Registros"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Seleccion 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1020
      Width           =   6570
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6570
      Begin VB.CheckBox chkImpFecha 
         Caption         =   "Imprime Fecha"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4740
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox Activar 
         Alignment       =   1  'Right Justify
         Caption         =   "Rango Periodo"
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   0
         Width           =   1425
      End
      Begin VB.ComboBox cboTpoMon 
         Height          =   315
         ItemData        =   "frmLibros.frx":0000
         Left            =   2565
         List            =   "frmLibros.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbEjercicio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   2565
         TabIndex        =   6
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Mes"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar toolbar 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   1005
      ButtonWidth     =   2434
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Libros y Registros"
            Object.ToolTipText     =   "Exportar Registro de Documentos a Excel"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   32
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A2"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A3"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A4"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A5"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A6"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A7"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A8"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A9"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A10"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A11"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A12"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A13"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A14"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A15"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A16"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A17"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A18"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A19"
               EndProperty
               BeginProperty ButtonMenu20 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A20"
               EndProperty
               BeginProperty ButtonMenu21 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A21"
               EndProperty
               BeginProperty ButtonMenu22 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A22"
               EndProperty
               BeginProperty ButtonMenu23 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A23"
               EndProperty
               BeginProperty ButtonMenu24 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A24"
               EndProperty
               BeginProperty ButtonMenu25 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A25"
               EndProperty
               BeginProperty ButtonMenu26 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A26"
               EndProperty
               BeginProperty ButtonMenu27 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A27"
               EndProperty
               BeginProperty ButtonMenu28 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A28"
               EndProperty
               BeginProperty ButtonMenu29 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A29"
               EndProperty
               BeginProperty ButtonMenu30 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A30"
               EndProperty
               BeginProperty ButtonMenu31 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A31"
               EndProperty
               BeginProperty ButtonMenu32 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A32"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vista Preliminar"
            Key             =   "B1"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PDT-Electrónicos"
            Key             =   "C1"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "D1"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4965
         Top             =   0
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
               Picture         =   "frmLibros.frx":0004
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibros.frx":015E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibros.frx":02B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibros.frx":0982
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   3435
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'23/03/2015 convirtiendo valores de mysql en sql
'Option Explicit 'se puso validar las porciones que se han sacaro de archivoplano
  
Public udFecha As Date
Public unCopias As Integer
Public unMargenIzquierdo As Integer
Public usDEstino As String
Public usOrientacionRpt As String
Public usOrientacionOri As String
Private pocnnMain As ADODB.Connection
Private porstMRp As ADODB.Recordset
Private pnOpcion As Integer

Private s_Caracter As String '2016-02-02.03 correccion ple

Private Sub Activar_Click()

If Activar.Value = Checked Then
    cmbEjercicio.Enabled = True
    cboTpoMon.Enabled = True
    chkImpFecha.Enabled = True
Else
    cmbEjercicio.Enabled = False
    cboTpoMon.Enabled = False
    chkImpFecha.Enabled = False
End If

End Sub

Private Sub Form_Load()
Dim i As Integer
 
pnOpcion = "01"

With cboTpoMon
  .AddItem TPOMON_NAC_TXT_1, 0
  .AddItem TPOMON_EXT_TXT_1, 1
End With

cboTpoMon.ListIndex = 0

cmbEjercicio.AddItem "00 Apertura"
cmbEjercicio.AddItem "01 Enero"
cmbEjercicio.AddItem "02 Febrero"
cmbEjercicio.AddItem "03 Marzo"
cmbEjercicio.AddItem "04 Abril"
cmbEjercicio.AddItem "05 Mayo"
cmbEjercicio.AddItem "06 Junio"
cmbEjercicio.AddItem "07 Julio"
cmbEjercicio.AddItem "08 Agosto"
cmbEjercicio.AddItem "09 Septiembre"
cmbEjercicio.AddItem "10 Octubre"
cmbEjercicio.AddItem "11 Noviembre"
cmbEjercicio.AddItem "12 Diciembre"
cmbEjercicio.AddItem "13 Cierre"

Select Case gsMesAct
Case "00"
    cmbEjercicio.Text = "00 Apertura"
Case "01"
    cmbEjercicio.Text = "01 Enero"
Case "02"
    cmbEjercicio.Text = "02 Febrero"
Case "03"
    cmbEjercicio.Text = "03 Marzo"
Case "04"
    cmbEjercicio.Text = "04 Abril"
Case "05"
    cmbEjercicio.Text = "05 Mayo"
Case "06"
    cmbEjercicio.Text = "06 Junio"
Case "07"
    cmbEjercicio.Text = "07 Julio"
Case "08"
    cmbEjercicio.Text = "08 Agosto"
Case "09"
    cmbEjercicio.Text = "09 Septiembre"
Case "10"
    cmbEjercicio.Text = "10 Octubre"
Case "11"
    cmbEjercicio.Text = "11 Noviembre"
Case "12"
    cmbEjercicio.Text = "12 Diciembre"
Case "13"
    cmbEjercicio.Text = "13 Cierre"
End Select

xqmes = gsMesAct

toolbar.Buttons(1).ButtonMenus(1).Text = "Formato 1.1  : Libro Caja y Bancos Detalle de los Movimientos del Efectivo"
toolbar.Buttons(1).ButtonMenus(2).Text = "Formato 1.2  : Libro Caja y Bancos Detalle de los Movimientos de la Cuenta Corriente"
toolbar.Buttons(1).ButtonMenus(3).Text = "-"
toolbar.Buttons(1).ButtonMenus(4).Text = "Formato 3.1  : Libro de Inventarios y Balances - Balance General (Configuracion de Estados Financieros)"
toolbar.Buttons(1).ButtonMenus(5).Text = "Formato 3.2  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 10 Caja y Bancos"
toolbar.Buttons(1).ButtonMenus(6).Text = "Formato 3.3  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 12 Cuentas Por Cobrar Clientes"
toolbar.Buttons(1).ButtonMenus(7).Text = "Formato 3.3  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 13 Cuentas Por Cobrar Relacionadas"
toolbar.Buttons(1).ButtonMenus(8).Text = "Formato 3.4  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 14 Cuentas Por Cobrar a Accionistas (o Socios) y Personal"
toolbar.Buttons(1).ButtonMenus(9).Text = "Formato 3.5  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 16 Cuentas Por Cobrar Diversas"
toolbar.Buttons(1).ButtonMenus(10).Text = "Formato 3.5  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 17 Cuentas Por Cobrar Relacionadas"
toolbar.Buttons(1).ButtonMenus(11).Text = "Formato 3.6  : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 19 Provisión por Cuenta de Cobranza Dudosa"
toolbar.Buttons(1).ButtonMenus(12).Text = "Formato 3.11 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 41 Remuneraciones por Pagar"
toolbar.Buttons(1).ButtonMenus(13).Text = "Formato 3.12 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 42 Proveedores"
toolbar.Buttons(1).ButtonMenus(14).Text = "Formato 3.12 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 43 Proveedores - Relacionadas"
toolbar.Buttons(1).ButtonMenus(15).Text = "Formato 3.13 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 46 Cuentas Por Pagar Diversas"
toolbar.Buttons(1).ButtonMenus(16).Text = "Formato 3.13 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 47 Cuentas Por Pagar Diversas - Relacionadas"
toolbar.Buttons(1).ButtonMenus(17).Text = "Formato 3.14 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 47 Beneficios Sociales de los Trabajadores"
toolbar.Buttons(1).ButtonMenus(18).Text = "Formato 3.15 : Libro de Inventarios y Balances - Detalle del Saldo de la Cuenta 49 Ganancias Diferidas"
toolbar.Buttons(1).ButtonMenus(19).Text = "Formato 3.17 : Libro de Inventarios y Balances - Balance de Comprobacion"
toolbar.Buttons(1).ButtonMenus(20).Text = "Formato 3.18 : Libro de Inventarios y Balances - Estados de Flujos de Efectivo"
toolbar.Buttons(1).ButtonMenus(21).Text = "Formato 3.20 : Libro de Inventarios y Balances - Estado de Ganancias y Perdidas por Funcion al 01.01 al 31.12 (Configuracion de Estados Financieros)"
toolbar.Buttons(1).ButtonMenus(22).Text = "-"
toolbar.Buttons(1).ButtonMenus(23).Text = "Formato 5.1  : Libro Diario"

toolbar.Buttons(1).ButtonMenus(24).Text = "Formato 5.2  : Libro Diario de Formato Simplificado"
toolbar.Buttons(1).ButtonMenus(25).Text = "Formato 5.2  : Libro Diario de Formato Simplificado Resumen"

toolbar.Buttons(1).ButtonMenus(26).Text = "-"
toolbar.Buttons(1).ButtonMenus(27).Text = "Formato 6.1  : Libro Mayor"
toolbar.Buttons(1).ButtonMenus(28).Text = "-"
toolbar.Buttons(1).ButtonMenus(29).Text = "Formato 8.1.1: Registro de Compras (A3)"
toolbar.Buttons(1).ButtonMenus(30).Text = "Formato 8.1.2: Registro de Compras (A4)"
'toolbar.Buttons(1).ButtonMenus(30).Text = "-"
toolbar.Buttons(1).ButtonMenus(31).Text = "Formato 14.1 : Registro de Ventas"

'ini 2014-05-30 adicion 5.3 plan ctas
toolbar.Buttons(1).ButtonMenus(32).Text = "Formato 5.3  : Detalle Plan de Cuenta Utilizado"
'fin 2014-05-30 adicion 5.3 plan ctas


'toolbar.Buttons(1).ButtonMenus(27).Text = "-"
'toolbar.Buttons(1).ButtonMenus(28).Text = "Formato 5.2  : Libro Diario de Formato Simplificado Resumen"

'ini 2015-04-21 nuevo reporte balance
'toolbar.Buttons(1).ButtonMenus(33).Text = "-"
'toolbar.Buttons(1).ButtonMenus(34).Text = "Formato 15.1  : 3.2 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 10 EFECTIVO Y EQUIVALENTES DE EFECTIVO (PCGE) (2)"
'fin 2015-04-21 nuevo reporte balance

'Deshabilitar Toolbars
toolbar.Buttons(1).ButtonMenus(4).Enabled = False
toolbar.Buttons(1).ButtonMenus(20).Enabled = False
toolbar.Buttons(1).ButtonMenus(21).Enabled = False

toolbar.Buttons(1).ButtonMenus(17).Enabled = False

Seleccion.Text = toolbar.Buttons(1).ButtonMenus(1).Text

Set pocnnMain = New ADODB.Connection
Set porstMRp = New ADODB.Recordset
   
'On Error GoTo Err
 
With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
End With

With porstMRp
      .ActiveConnection = pocnnMain
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
End With

'Características de impresión.
    udFecha = Date
    unCopias = 1
    unMargenIzquierdo = 240
    usDEstino = PRN_DEST_MATR
    usOrientacionRpt = PRN_ORIE_VERT
']

frmOPrnCfg.OrientacionPrn 0, Me
frmOPrnCfg.lblOriPrn.Caption = Printer.Orientation
   
cmbEjercicio.Enabled = False
cboTpoMon.Enabled = False
chkImpFecha.Enabled = False
  
Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub




Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
   Case "B1": Imprimir
   Case "C1": ppRegistroElectronico
   Case "D1": Unload Me
  End Select
End Sub

Private Sub toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Seleccion.Text = ButtonMenu.Text
  Select Case ButtonMenu.Key
   Case "A" & Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
    pnOpcion = Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
  End Select
End Sub
Sub BalancedeComprobacion()
  Dim dnContador As Integer, n_Index As Integer
  Dim s_Sentencia As String, s_Sql As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_Moneda As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_MesIni As Integer, n_MesFin As Integer
  Dim s_Expresion As String
  Dim l_CreateTB As Boolean

  s_AnoIni = gsAnoAct
  s_AnoFin = gsAnoAct
    
  'CORREGIR NIVEL DE CUENTA
  s_Expresion = gsNivCta
  pnNivCta = Right(s_Expresion, 1)
  s_Moneda = TPOMON_NAC_TXT
   
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
  
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    n_MesIni = 1
    n_MesFin = Val(gsMesAct)
    ' Acumulación de saldos
    s_SaldoDeb = "ROUND(0": s_SaldoHab = "ROUND(0"
    If gsMesAct <> "00" Then
      s_SaldoDeb = "ROUND(": s_SaldoHab = "ROUND("
      For n_Index = n_MesIni To n_MesFin
        s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
        s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      Next n_Index
    End If
    s_SaldoDeb = s_SaldoDeb & ", 2)"
    s_SaldoHab = s_SaldoHab & ", 2)"
    
       
    ' Registros iniciales de saldos
    s_Sentencia = "SELECT a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, b.TpoSdo, b.TpoCta, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & " AS cSumaD, " & s_SaldoHab & " AS cSumaH, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoDeb & " ELSE 0 END) AS cSumaDt, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoHab & " ELSE 0 END) AS cSumaHt,round(A.AcuD00_" & s_Moneda & ",2) AS cApeD, round(A.AcuH00_" & s_Moneda & ",2) AS cApeH "
        
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCpb ", "")
    s_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    If pnNivCta = 2 Then
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    Else
      s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
      s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<" & pnNivCta & " AND b.TpoCta=" & TPOCTA_TRA & ")) "
    End If
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "HAVING (cSumaD + cSumaH + cSumaDt + cSumaHt + cApeD + cApeH) > 0 "
    End If
    ' Registros iniciales
    's_Sentencia = s_Sentencia & " UNION ALL "
    's_Sentencia = s_Sentencia & " SELECT a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, 'X', b.TpoCta, "
    's_Sentencia = s_Sentencia & " a.AcuD00_" & s_Moneda & " AS cSumaD, a.AcuH00_" & s_Moneda & " AS cSumaH, "
    's_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    's_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    's_Sentencia = s_Sentencia & "THEN a.AcuD00_" & s_Moneda & " ELSE 0 END) AS cSumaDt, "
    's_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    's_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    's_Sentencia = s_Sentencia & "THEN a.AcuH00_" & s_Moneda & " ELSE 0 END) AS cSumaHt "
    's_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCpb ", "")
    's_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
    's_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    's_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    's_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    'If pnNivCta = 2 Then
    '  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    'Else
    '  s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    '  s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<" & pnNivCta & " AND b.TpoCta=" & TPOCTA_TRA & ")) "
    'End If
    'If ps_Plataforma = pSrvMySql Then
    '  s_Sentencia = s_Sentencia & "HAVING (cSumaD + cSumaH + cSumaDt + cSumaHt) > 0 "
    'End If
    s_Sentencia = s_Sentencia & "ORDER BY 1"
    ' Executo la sentencia
    If Not l_CreateTB Then
      's_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trpRngBceCpb ", "")
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS trpRngBceCpb ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trpRngBceCpb "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
   
  With porstMRp
    If .State = adStateOpen Then .Close
    s_Sentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(cApeD), 2) AS cApeD, ROUND(SUM(cApeH), 2) AS cApeH "
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
    s_Sentencia = s_Sentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
    s_Sentencia = s_Sentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)+ ROUND(SUM(cApeD), 2)+ ROUND(SUM(cApeH), 2)) > 0 "
    s_Sentencia = s_Sentencia & "ORDER BY CodCta"
    .Source = s_Sentencia
    .Open
  End With
  
  s_Sentencia = ""
  gpEncabezadoRpt frmMain.rptMain, "Formato 3.17: Balance de Comprobación" & " (" & TPOMON_NAC_TXT_1 & " )", udFecha, True, 0, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "LOBalComp.rpt"
      .ParameterFields(1) = "Fecha;" & IIf(chkImpFecha.Value = Checked, 1, 0) & ";true"
      .ParameterFields(2) = "RepLegal;" & gsRepEmp & ";true"
      .ParameterFields(3) = "Contador;" & gsConEmp & ";true"

      'Fórmular propias.
      .Formulas(5) = "mPeriodo='" & s_Sentencia & " " & Choose(gsIdioma, "Acumulado - ", "Accrued - ") & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
      .WindowShowExportBtn = True
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  ' elimino el archivo temporal
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")

End Sub
'ini 2015-12-10 PLE rpt fmt electro archivo

Sub BalancedeComprobacion2()
  Dim dnContador As Integer, n_Index As Integer
  Dim s_Sentencia As String, s_Sql As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_Ano As String, s_Mes As String
  Dim s_Moneda As String, s_Catalogo As String
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim n_MesIni As Integer, n_MesFin As Integer
  Dim s_Expresion As String
  Dim l_CreateTB As Boolean

  s_AnoIni = gsAnoAct
  s_AnoFin = gsAnoAct
    
  'CORREGIR NIVEL DE CUENTA
  s_Expresion = gsNivCta
  pnNivCta = Right(s_Expresion, 1)
  s_Moneda = TPOMON_NAC_TXT
   
  ' Elimino y genero el archivo temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
  
  For dnContador = Val(s_AnoIni) To Val(s_AnoFin)
    s_Ano = Trim$(dnContador)
    s_Catalogo = s_Ano
    n_MesIni = 1
    n_MesFin = Val(gsMesAct)
    ' Acumulación de saldos
    s_SaldoDeb = "ROUND(0": s_SaldoHab = "ROUND(0"
    If gsMesAct <> "00" Then
      s_SaldoDeb = "ROUND(": s_SaldoHab = "ROUND("
      For n_Index = n_MesIni To n_MesFin
        s_SaldoDeb = s_SaldoDeb & "a.AcuD" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
        s_SaldoHab = s_SaldoHab & "a.AcuH" & Format(Trim(n_Index), "00") & "_" & s_Moneda & IIf(n_Index = n_MesFin, "", "+")
      Next n_Index
    End If
    s_SaldoDeb = s_SaldoDeb & ", 2)"
    s_SaldoHab = s_SaldoHab & ", 2)"
    
       
    ' Registros iniciales de saldos
    s_Sentencia = "SELECT a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, b.TpoSdo, b.TpoCta, "
    s_Sentencia = s_Sentencia & s_SaldoDeb & " AS cSumaD, " & s_SaldoHab & " AS cSumaH, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoDeb & " ELSE 0 END) AS cSumaDt, "
    s_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    s_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    s_Sentencia = s_Sentencia & "THEN " & s_SaldoHab & " ELSE 0 END) AS cSumaHt,round(A.AcuD00_" & s_Moneda & ",2) AS cApeD, round(A.AcuH00_" & s_Moneda & ",2) AS cApeH "
        
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCpb ", "")
    s_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
    s_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    If pnNivCta = 2 Then
      s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    Else
      s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
      s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<" & pnNivCta & " AND b.TpoCta=" & TPOCTA_TRA & ")) "
    End If
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "HAVING (cSumaD + cSumaH + cSumaDt + cSumaHt + cApeD + cApeH) > 0 "
    End If
    ' Registros iniciales
    's_Sentencia = s_Sentencia & " UNION ALL "
    's_Sentencia = s_Sentencia & " SELECT a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, 'X', b.TpoCta, "
    's_Sentencia = s_Sentencia & " a.AcuD00_" & s_Moneda & " AS cSumaD, a.AcuH00_" & s_Moneda & " AS cSumaH, "
    's_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    's_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    's_Sentencia = s_Sentencia & "THEN a.AcuD00_" & s_Moneda & " ELSE 0 END) AS cSumaDt, "
    's_Sentencia = s_Sentencia & "(CASE WHEN ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & ") "
    's_Sentencia = s_Sentencia & "OR ((" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<=" & pnNivCta & ") AND b.TpoCta='" & TPOCTA_TRA & "')) "
    's_Sentencia = s_Sentencia & "THEN a.AcuH00_" & s_Moneda & " ELSE 0 END) AS cSumaHt "
    's_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql And Not l_CreateTB, "INTO #trpRngBceCpb ", "")
    's_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
    's_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
    's_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    's_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
    'If pnNivCta = 2 Then
    '  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    'Else
    '  s_Sentencia = s_Sentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))=" & pnNivCta & " "
    '  s_Sentencia = s_Sentencia & "OR (" & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCta))<" & pnNivCta & " AND b.TpoCta=" & TPOCTA_TRA & ")) "
    'End If
    'If ps_Plataforma = pSrvMySql Then
    '  s_Sentencia = s_Sentencia & "HAVING (cSumaD + cSumaH + cSumaDt + cSumaHt) > 0 "
    'End If
    s_Sentencia = s_Sentencia & "ORDER BY 1"
    ' Executo la sentencia
    If Not l_CreateTB Then
      's_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trpRngBceCpb ", "")
      s_Sql = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS trpRngBceCpb ", "")
      l_CreateTB = True
    Else
      s_Sql = "INSERT INTO " & ps_Prefijo & "trpRngBceCpb "
    End If
    s_Sql = s_Sql & s_Sentencia
    pocnnMain.Execute s_Sql
  Next dnContador
   
'''  With porstMRp
'''    If .State = adStateOpen Then .Close
'''    s_Sentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
'''    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
'''    s_Sentencia = s_Sentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt, "
'''    s_Sentencia = s_Sentencia & "ROUND(SUM(cApeD), 2) AS cApeD, ROUND(SUM(cApeH), 2) AS cApeH "
'''    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
'''    s_Sentencia = s_Sentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
'''    s_Sentencia = s_Sentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)+ ROUND(SUM(cApeD), 2)+ ROUND(SUM(cApeH), 2)) > 0 "
'''    s_Sentencia = s_Sentencia & "ORDER BY CodCta"
'''    .Source = s_Sentencia
'''    .Open
'''  End With
'
'''  s_Sentencia = ""
'''  gpEncabezadoRpt frmMain.rptMain, "Formato 3.17: Balance de Comprobación" & " (" & TPOMON_NAC_TXT_1 & " )", udFecha, True, 0, porstMRp
'''    With frmMain.rptMain
'''      '[Datos y parámetros del reporte.  'Cambiar.
'''      .ReportFileName = gsRutRpt & "LOBalComp.rpt"
'''      .ParameterFields(1) = "Fecha;" & IIf(chkImpFecha.Value = Checked, 1, 0) & ";true"
'''      .ParameterFields(2) = "RepLegal;" & gsRepEmp & ";true"
'''      .ParameterFields(3) = "Contador;" & gsConEmp & ";true"
'''
'''      'Fórmular propias.
'''      .Formulas(5) = "mPeriodo='" & s_Sentencia & " " & Choose(gsIdioma, "Acumulado - ", "Accrued - ") & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
'''      .WindowShowExportBtn = True
'''      .MarginLeft = unMargenIzquierdo
'''      .WindowState = crptMaximized
'''      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
'''      .Action = 1
'''    End With
'''  ' elimino el archivo temporal
'''  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")

End Sub
'fin 2015-12-10 PLE rpt fmt electro archivo

Private Sub Imprimir()
  Dim sql As String, sReporte As String, sMoneda As String
  Dim i As Integer
  Dim valid As Boolean

  valid = False
  If pnOpcion = 99 Then MsgBox "Seleccionar Libro o Registro", vbCritical, "Sistema Contable": Exit Sub
  If cboTpoMon.Text = "" Then MsgBox "Seleccionar Moneda", vbCritical, "Sistema Contable": Exit Sub
  
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "MN", "ME")
  xqmes = Left(cmbEjercicio.Text, 2)
  
  rTitulo = ""
  'ini rcs 2015-03-23 correccion version
        Dim arrDocume(4) As String
        arrDocume(0) = "d.AbvTDc"
        arrDocume(1) = "'-'"
        arrDocume(2) = "a.SerDoc"
        arrDocume(3) = "'-'"
        arrDocume(4) = "a.NroDoc"
        
        'FORMATO 8.1.1 y 8.1.2
        Dim arrnrocpb(1) As String
        arrnrocpb(0) = "com.coddro"
        arrnrocpb(1) = "com.nrocpb"

        'FORMATO 8.1.1 y 8.1.2
        Dim arrv1(2) As String
        arrv1(0) = "serdoc_ref"
        arrv1(1) = "'-'"
        arrv1(2) = "nrodoc_ref"
  'fin rcs 2015-03-23 correccion version
  
  Select Case pnOpcion
   Case 1: LibroCaja 101: Exit Sub
   Case 2: LibroCaja 104: Exit Sub
   Case 5: LibroCaja 999: Exit Sub
   Case 6
    sReporte = "LOInBal.rpt"
    'rTitulo = "FORMATO 3.3: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 12 CUENTAS POR COBRAR COMERCIALES - TERCEROS "
    rTitulo = "FORMATO 3.3: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 12 CUENTAS POR COBRAR CLIENTES "
    'ini rcs 2015-03-23 correccion version sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite ,MAX(a.refdoc) refdoc ,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    sql = sql & " AND left(a.CodCta,2)='12' "
    'sql = sql & " AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 "
    sql = sql & " and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 correccion version sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING "
        sql = sql & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
        sql = sql & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    
   Case 7
    sReporte = "LOInBal.rpt"
    'rTitulo = "FORMATO 3.3: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 13 CUENTAS POR COBRAR COMERCIALES - RELACIONADAS"
    rTitulo = "FORMATO 3.3: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 13 CUENTAS POR COBRAR - RELACIONADAS"
    'sql8 2012-04-15 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite ,MAX(a.refdoc) refdoc ,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    sql = sql & " AND left(a.CodCta,2)='13' "
    'sql = sql & " AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 "
    sql = sql & " and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING "
        sql = sql & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
        sql = sql & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 8
    sReporte = "LOInBal.rpt"
    'rTitulo = "FORMATO 3.4: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS Y DIRECTOR"
    rTitulo = "FORMATO 3.4: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR A ACCIONISTAS (O SOCIOS) Y PERSONAL"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci), c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sql = sql & " AND left(a.CodCta,2)='14' AND IFNULL(a.CodAux, '') <>''"
    'sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='14' "
    sql = sql & " AND mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
        sql = sql & "      OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
     sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 9
    sReporte = "LOInBal.rpt"
    'rTitulo = "FORMATO 3.5: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS-TERCEROS"
    rTitulo = "FORMATO 3.5: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    sql = sql & " AND left(a.CodCta,2)='16'  AND IFNULL(a.CodAux, '') <>''"
    sql = sql & " AND mespvs <= " & Left(cmbEjercicio.Text, 2)
    'sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING "
        sql = sql & "  (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
        sql = sql & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 10
    sReporte = "LOInBal.rpt"
    'rTitulo = "FORMATO 3.5: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 17 CUENTAS POR COBRAR DIVERSAS - RELACIONADAS"
    rTitulo = "FORMATO 3.5: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 17 CUENTAS POR COBRAR RELACIONADAS"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & "  AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sql8 2015-03-23 sql = sql & " AND left(a.CodCta,2)='17'  AND IFNULL(a.CodAux, '') <>''"
    sql = sql & " AND left(a.CodCta,2)='17'  AND " & fIsNull() & "a.CodAux, '') <>''"
    'sql8 2015-03-23 sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING "
        sql = sql & " (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "  ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 11
    sReporte = "LOInBal.rpt"
    'rTitulo = "FORMATO 3.6: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA"
    rTitulo = "FORMATO 3.6: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 19 PROVISION PARA CUENTA DE COBRANZA DUDOSA"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda, MAX(a.gloite)gloite, MAX(a.refdoc) refdoc, MAX(a.coddro) coddro, MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sql8 2015-03-23 sql = sql & " AND left(a.CodCta,2)='19' AND IFNULL(a.CodAux, '') <>''"
    'sql8 2015-03-23 sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='19' AND " & fIsNull() & "a.CodAux, '') <>''"
    sql = sql & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
         sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING "
        sql = sql & " (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "  ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 12
    sReporte = "LOInBalP.rpt"
    'rTitulo = "FORMATO 3.11: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 41 REMUNERACIONES Y PARTICIPACIONES POR PAGAR"
    rTitulo = "FORMATO 3.11: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 41 REMUNERACIONES POR PAGAR"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sql = sql & " AND left(a.CodCta,2)='41' AND IFNULL(a.CodAux, '') <>''"
    'sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='41' "
    sql = sql & " and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING "
        sql = sql & " (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "  ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
        sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 13
    sReporte = "LOInBalP.rpt"
    'rTitulo = "FORMATO 3.12: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES - TERCEROS"
    rTitulo = "FORMATO 3.12: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 42 PROVEEDORES"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sql8 2015-03-23 sql = sql & " AND left(a.CodCta,2)='42' AND IFNULL(a.CodAux, '') <>''"
    'sql8 2015-03-23 sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='42' AND " & fIsNull() & "a.CodAux, '') <>''"
    sql = sql & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sql8 2015-03-23 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "        ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & "  ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 14
    sReporte = "LOInBalP.rpt"
    'rTitulo = "FORMATO 3.12: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 43 CUENTAS POR PAGAR COMERCIALES - RELACIONADAS"
    rTitulo = "FORMATO 3.12: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 43 PROVEEDORES RELACIONADAS"
    '2015-03-25 mysql a sql8 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
'ini 2015-03-25 mysql a sql8
'    sql = sql & " AND left(a.CodCta,2)='43' AND IFNULL(a.CodAux, '') <>''"
'    sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='43' AND " & fIsNull() & "a.CodAux, '') <>''"
    sql = sql & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
'fin 2015-03-25 mysql a sql8
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    '2015-03-25 mysql a sql8 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
       '2015-06-11 error order by sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
       sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "        ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
     sql = sql & "  ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 15
    sReporte = "LOInBalP.rpt"
    'rTitulo = "FORMATO 3.13: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 46 CUENTAS POR PAGAR DIVERSAS - TERCEROS"
    rTitulo = "FORMATO 3.13: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 46 CUENTAS POR PAGAR DIVERSAS"
    '2015-03-25 mysql a sql8 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    '2015-03-25 mysql a sql8 sql = sql & " AND left(a.CodCta,2)='46' AND IFNULL(a.CodAux, '') <>''"
    sql = sql & " AND left(a.CodCta,2)='46' AND " & fIsNull() & "a.CodAux, '') <>''"
    sql = sql & " AND mespvs <= " & Left(cmbEjercicio.Text, 2)
    'sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    '2015-03-25 mysql a sql8 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "        ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 16
    sReporte = "LOInBalP.rpt"
    'rTitulo = "FORMATO 3.13: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 47 CUENTAS POR PAGAR DIVERSAS - RELACIONADAS"
    rTitulo = "FORMATO 3.13: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 47 CUENTAS POR PAGAR DIVERSAS - RELACIONADAS"
    '2015-03-25 mysql a sql8 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    '2015-03-25 mysql a sql8 sql = sql & " AND left(a.CodCta,2)='47' AND IFNULL(a.CodAux, '') <>''"
    '2015-03-25 mysql a sql8 sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='47' AND " & fIsNull() & "a.CodAux, '') <>''"
    sql = sql & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    '2015-03-25 mysql a sql8 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "        ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 17
    sReporte = "LOInBalP.rpt"
    rTitulo = "FORMATO 3.14: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 47 BENEFICIOS SOCIALES DE LOS TRABAJADORES"
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    sql = sql & " AND left(a.CodCta,2)='47' AND IFNULL(a.CodAux, '') <>''"
    sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 18
    sReporte = "LOInBalP.rpt"
    'rTitulo = "FORMATO 3.15: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 49 PASIVO POR IMPUESTO A LA RENTA Y PARTICIPACIONES DE LOS TRAB"
    rTitulo = "FORMATO 3.15: LIBRO DE INVENTARIOS Y BALANCES DETALLE DEL SALDO DE LA CUENTA 49 GANANCIAS DIFERIDAS"
    '2015-03-25 mysql a sql8 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sql = sql & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sql = sql & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sql = sql & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sql = sql & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
'ini 2015-03-25 mysql a sql8
'    sql = sql & " AND left(a.CodCta,2)='49' AND IFNULL(a.CodAux, '') <>''"
'    sql = sql & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sql = sql & " AND left(a.CodCta,2)='49' AND " & fIsNull() & "a.CodAux, '') <>''"
    sql = sql & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
'fin 2015-03-25 mysql a sql8
    sql = sql & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    '2015-03-25 mysql a sql8 sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sql = sql & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sql = sql & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sql = sql & "        ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sql = sql & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   Case 19: BalancedeComprobacion: Exit Sub
   Case 23
    sReporte = "LODiario.rpt"
    rTitulo = "FORMATO 5.1: LIBRO DIARIO"
    sql = "SELECT LEFT(a.CodDro,2) AS cDiario, a.CodDro, a.NroCpb, a.FehOpe, a.CodTDc, a.SerDoc, a.NroDoc, "
    sql = sql & " a.CodCta, f.detcta,a.CodAux, b.RazAux, a.RefDoc, " & Choose(gsIdioma, "left(a.GloIte,30)", "left(a.GloItex,30)") & " AS GloIte, a.TpoCtb, "
    sql = sql & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    sql = sql & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodAux, '-', b. RazAux)", "(a.CodAux+'-'+b. RazAux)") & " AS cx1, "
    sql = sql & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    sql = sql & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    sql = sql & "c.AbvTDc, " & Choose(gsIdioma, "e.DetDro", "e.DetDrox") & " AS DetDro, "
    sql = sql & Choose(gsIdioma, "d.DetDro", "d.DetDrox") & " AS cDetSubDro,'" & sMoneda & "' as Moneda,d.codlib "
    sql = sql & "FROM ((((COCpbDet a "
    sql = sql & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    sql = sql & "LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    sql = sql & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    sql = sql & "LEFT JOIN CODro e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND LEFT(a.CodDro, 2)=RTrim(e.CodDro)) "
    sql = sql & "LEFT JOIN CoCta f ON a.codemp=f.codemp AND a.pdoano=f.pdoano AND a.Codcta=f.Codcta "
    sql = sql & "WHERE a.codemp='" & gsCodEmp & "' "
    sql = sql & "AND a.pdoano='" & gsAnoAct & "' "
    sql = sql & "AND a.Mespvs ='" & xqmes & "' "
    sql = sql & "AND NOT(a.tpognr='" & TPOGNR_DCA & "' AND a.imp" & sMoneda & "=0.00) "
    sql = sql & "ORDER BY a.CodDro, a.NroCpb, a.FehOpe "
   
   Case 24
    sReporte = "LODiarioSimplificado.rpt"
    rTitulo = "FORMATO 5.2: LIBRO DIARIO DE FORMATO SIMPLIFICADO"
    '2015-03-25 mysql a sql8 sql = "select concat(cocpbdet.coddro,'-',nrocpb),"
    Dim arrConcat(2) As String
    arrConcat(0) = "cocpbdet.coddro"
    arrConcat(1) = "'-'"
    arrConcat(2) = "nrocpb"
    sql = "select " & fConCat(arrConcat) & ","
    sql = sql & " fehope,"
    sql = sql & " cocpbdet.coddro,"
'2015-06-16 error duplic    sql = sql & " fehope,"
'2015-06-16 error duplic    sql = sql & " cocpbdet.coddro,"
    '2015-03-25 mysql a sql8 sql = sql & " concat(left(IFNULL(gloite, ''), 22), (CASE WHEN ISNULL(abvtdc) THEN '' ELSE ' - ' END), IFNULL(abvtdc, ''), (CASE WHEN ISNULL(serdoc) THEN '' ELSE ' - ' END), IFNULL(serdoc, ''),'-',IFNULL(nrodoc, '')) as detdro,  codcta as Cuenta,"
    Dim arrdetdro(6) As String
    arrdetdro(0) = "left(" & fIsNull() & "gloite, ''), 22)"
    '2015-04-09 error null arrdetdro(1) = "(CASE ISNULL(abvtdc,'') WHEN '' THEN '' ELSE ' - ' END)"
    arrdetdro(1) = "(CASE " & fIsNull() & "abvtdc,'') WHEN '' THEN '' ELSE ' - ' END)"
    arrdetdro(2) = "" & fIsNull() & "abvtdc, '')"
    '2015-04-09 error null arrdetdro(3) = "(CASE ISNULL(serdoc,'') WHEN '' THEN '' ELSE ' - ' END)"
    arrdetdro(3) = "(CASE " & fIsNull() & "serdoc,'') WHEN '' THEN '' ELSE ' - ' END)"
    arrdetdro(4) = "" & fIsNull() & "serdoc, '')"
    arrdetdro(5) = "'-'"
    arrdetdro(6) = "" & fIsNull() & "nrodoc, '')"
    sql = sql & " " & fConCat(arrdetdro) & " as detdro,"
    sql = sql & "   codcta as Cuenta,"
    sql = sql & " (case tpoctb when 'D' then imp" & sMoneda & " else 0 end)-(case tpoctb when 'H' then imp" & sMoneda & " else 0 end)"
    sql = sql & " from cocpbdet"
    sql = sql & " inner join codro on cocpbdet.coddro=codro.coddro and cocpbdet.codemp=codro.codemp and cocpbdet.pdoano=codro.pdoano"
    sql = sql & " left join tgtdc on cocpbdet.codemp=tgtdc.codemp and cocpbdet.codtdc=tgtdc.codtdc "
    sql = sql & " where cocpbdet.codemp='" & gsCodEmp & "' and cocpbdet.pdoano='" & gsAnoAct & "' and mespvs='" & xqmes & "'"
    sql = sql & " order by 1 "
   Case 25
    sReporte = "LODiarioSimplificadoR.rpt"
    rTitulo = "FORMATO 5.2: LIBRO DIARIO DE FORMATO SIMPLIFICADO RESUMEN"
'ini 2015-03-25 mysql a sql8
'    sql = "select '' as d1,'' as d2,'' as d3,cocta.detcta,cocpbdet.codcta,"
'    sql = sql & " sum((case tpoctb when 'D' then imp" & sMoneda & " else 0 end)) as debe,sum((case tpoctb when 'H' then imp" & sMoneda & " else 0 end)) as haber"
'    sql = sql & " from cocpbdet"
'    sql = sql & " inner join cocta on cocpbdet.codcta=cocta.codcta and cocpbdet.codemp=cocta.codemp and cocpbdet.pdoano=cocta.pdoano "
'    sql = sql & " where cocpbdet.codemp='" & gsCodEmp & "' and cocpbdet.pdoano='" & gsAnoAct & "' and mespvs='" & xqmes & "'"
'    sql = sql & " group by cocpbdet.codcta "
        sql = "select '' as d1,'' as d2,'' as d3,max(a.detcta) detcta,cocpbdet.codcta,"
        sql = sql & "    sum((case tpoctb when 'D' then impMN else 0 end)) as debe,"
        sql = sql & "    sum((case tpoctb when 'H' then impMN else 0 end)) as haber "
        sql = sql & "From cocpbdet "
        sql = sql & "inner join cocta a on cocpbdet.codcta=a.codcta and cocpbdet.codemp=a.codemp "
        sql = sql & "    and cocpbdet.pdoano=a.pdoano "
        sql = sql & "where cocpbdet.codemp='" & gsCodEmp & "' and cocpbdet.pdoano='" & gsAnoAct & "' and mespvs='" & xqmes & "' "
'2015-06-16 error where     sql = sql & "where cocpbdet.codemp='001' and cocpbdet.pdoano='2012' and mespvs='02' "
        sql = sql & "group by cocpbdet.codcta"
'fin 2015-03-25 mysql a sql8
   Case 27: LibroMayor: Exit Sub
   Case 29
    sReporte = "LOCompras.rpt"
    rTitulo = "FORMATO 8.1.1: REGISTRO DE COMPRAS"
    '2015-03-25 mysql a sql8  sql = "select concat(com.coddro,com.nrocpb) as nrocpb,date_format(com.feedoc,'%d/%m/%y') as feedoc,date_format(com.fevdoc,'%d/%m/%y') as fevdoc,com.codtdc as codtdc,"
    sql = "select " & fConCat(arrnrocpb) & " as nrocpb," & fConvert103ddmmyyySay("com.feedoc") & " as feedoc," & fConvert103ddmmyyySay("com.fevdoc") & " as fevdoc,com.codtdc as codtdc,"
    sql = sql & " case com.codtdc when '50' then codaduana when '52' then codaduana when '53' then codaduana else com.serdoc end as serdoc, com.annodua as anno,case com.codtdc when '50' then nrodua when '52' then nrodua when '53' then nrodua else com.nrodoc end as nrodoc, right(aux.tpodci,1) as tpodci,com.codaux as codaux,aux.razaux as razaux,"
    sql = sql & " impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp1,"
    sql = sql & " impigv_ogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp2,"
    sql = sql & " impogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp3,"
    sql = sql & " impigv_ogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp4,"
    sql = sql & " impong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp5,"
    sql = sql & " impigv_ong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp6,"
    sql = sql & " impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp7,"
    sql = sql & " impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp8,"
    sql = sql & " impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp9,"
    sql = sql & " imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp10,"
    '2015-03-25 mysql a sql8 sql = sql & " CASE WHEN codtdc_ref='91' THEN Concat(serdoc_ref, '-', nrodoc_ref) ELSE Null END as v1,"
    sql = sql & " CASE WHEN codtdc_ref='91' THEN " & fConCat(arrv1) & " ELSE Null END as v1,"
'ini 2015-09-16 t.cam solo ME
    'sql = sql & " nrocdt,fehcdt,imptcb,feedoc_ref,codtdc_ref,serdoc_ref,nrodoc_ref,'" & sMoneda & "' as Moneda "
    sql = sql & " nrocdt,fehcdt"
    sql = sql & " ,CASE tpomon WHEN '" & TPOMON_EXT & "' THEN imptcb ELSE 0 END imptcb"
    sql = sql & " ,feedoc_ref,codtdc_ref,serdoc_ref,nrodoc_ref,'" & sMoneda & "' as Moneda "
'fin 2015-09-16 t.cam solo ME
    sql = sql & " from cocprdoc com"
    sql = sql & " inner join tgaux aux on com.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "'"
    sql = sql & " inner join tgtdc tdc on com.codemp=tdc.codemp and com.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "'"
    sql = sql & " where com.codemp='" & gsCodEmp & "' and com.pdoano='" & gsAnoAct & "' and mespvs in (" & Left(cmbEjercicio.Text, 2) & ") order by 1"
   Case 30
    sReporte = "LOComprasa4.rpt"
    rTitulo = "FORMATO 8.1.2: REGISTRO DE COMPRAS A4"
    '2015-03-25 mysql a sql8 sql = "select concat(com.coddro,com.nrocpb) as nrocpb,date_format(com.feedoc,'%d/%m/%y') as feedoc,date_format(com.fevdoc,'%d/%m/%y') as fevdoc,com.codtdc as codtdc,"
    sql = "select " & fConCat(arrnrocpb) & " as nrocpb," & fConvert103ddmmyyySay("com.feedoc") & " as feedoc," & fConvert103ddmmyyySay("com.fevdoc") & " as fevdoc,com.codtdc as codtdc,"
    sql = sql & " case com.codtdc when '50' then codaduana when '52' then codaduana when '53' then codaduana else com.serdoc end as serdoc, com.annodua as anno,case com.codtdc when '50' then nrodua when '52' then nrodua when '53' then nrodua else com.nrodoc end as nrodoc, right(aux.tpodci,1) as tpodci,com.codaux as codaux,aux.razaux as razaux,"
    sql = sql & " impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp1,"
    sql = sql & " impigv_ogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp2,"
    sql = sql & " impogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp3,"
    sql = sql & " impigv_ogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp4,"
    sql = sql & " impong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp5,"
    sql = sql & " impigv_ong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp6,"
    sql = sql & " impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp7,"
    sql = sql & " impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp8,"
    sql = sql & " impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp9,"
    sql = sql & " imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp10,"
    '2015-03-25 mysql a sql8 sql = sql & " CASE WHEN codtdc_ref='91' THEN Concat(serdoc_ref, '-', nrodoc_ref) ELSE Null END as v1,"
    sql = sql & " CASE WHEN codtdc_ref='91' THEN " & fConCat(arrv1) & " ELSE Null END as v1,"
'ini 2015-09-16 t.cam solo ME
    'sql = sql & " nrocdt,fehcdt,imptcb,feedoc_ref,codtdc_ref,serdoc_ref,nrodoc_ref,'" & sMoneda & "' as Moneda "
    sql = sql & " nrocdt,fehcdt"
    sql = sql & " ,CASE tpomon WHEN '" & TPOMON_EXT & "' THEN imptcb ELSE 0 END imptcb"
    sql = sql & " ,feedoc_ref,codtdc_ref,serdoc_ref,nrodoc_ref,'" & sMoneda & "' as Moneda "
'fin 2015-09-16 t.cam solo ME
    sql = sql & " from cocprdoc com"
    sql = sql & " inner join tgaux aux on com.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "'"
    sql = sql & " inner join tgtdc tdc on com.codemp=tdc.codemp and com.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "'"
    sql = sql & " where com.codemp='" & gsCodEmp & "' and com.pdoano='" & gsAnoAct & "' and mespvs in (" & Left(cmbEjercicio.Text, 2) & ") order by 1"
   Case 31
    sReporte = "LOVentas.rpt"
    rTitulo = "FORMATO 14.1: REGISTRO DE VENTAS E INGRESOS"
    '2015-03-25 mysql a sql8 sql = "select concat(vta.coddro,vta.nrocpb) as nrocpb, date_format(vta.feedoc,'%d/%m/%y') as feedoc, date_format(vta.fevdoc,'%d/%m/%y') as fevdoc,vta.codtdc as codtdc,vta.serdoc as serdoc,"
    'FORMATO 8.1.1 y 8.1.2
        Dim arrnrocpb2(1) As String
        arrnrocpb2(0) = "vta.coddro"
        arrnrocpb2(1) = "vta.nrocpb"
    sql = "select " & fConCat(arrnrocpb2) & " as nrocpb," & fConvert103ddmmyyySay("vta.feedoc") & "as feedoc," & fConvert103ddmmyyySay("vta.fevdoc") & " as fevdoc,vta.codtdc as codtdc,vta.serdoc as serdoc,"
    
    '2015-03-25 mysql a sql8 sql = sql & " vta.nrodoc as nrodoc, aux.tpodci as tpodci, (CASE WHEN tpodci='01' THEN RIGHT(aux.rucaux, 8) ELSE aux.rucaux END) as codaux, trim(left(aux.razaux,60)) as razaux,"
    sql = sql & " vta.nrodoc as nrodoc,aux.tpodci as tpodci,vta.codaux as codaux," & fLTrim() & "left(aux.razaux,60)) as razaux,"
    sql = sql & " vta.impexp_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as impexp,"
    sql = sql & " vta.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as impogr,"
    sql = sql & " (CASE WHEN vta.categoriadoc<>'" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impexo,"
    sql = sql & " (CASE WHEN vta.categoriadoc='" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impina,"
    sql = sql & " vta.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as impisc,"
    sql = sql & " vta.impigv_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as impigv,"
    sql = sql & " vta.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as impoim,"
    sql = sql & " vta.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imptot,"
'ini 2015-09-16 t.cam solo ME
   'sql = sql & " vta.imptcb as imptcb,"
    sql = sql & " CASE tpomon WHEN '" & TPOMON_EXT & "' THEN vta.imptcb ELSE 0 END imptcb, "
'fin 2015-09-16 t.cam solo ME
    
    '2015-03-25 mysql a sql8 sql = sql & " date_format(feedoc_ref,'%d/%m/%y') as d1,codtdc_ref as d2,serdoc_ref as d3,nrodoc_ref as d4,'" & sMoneda & "' as Moneda "
    sql = sql & " " & fConvert103ddmmyyySay("feedoc_ref") & " as d1,codtdc_ref as d2,serdoc_ref as d3,nrodoc_ref as d4,'" & sMoneda & "' as Moneda "
    sql = sql & " from covtadoc vta  "
    sql = sql & " inner join tgaux aux on vta.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "'"
    sql = sql & " inner join tgtdc tdc on vta.codemp=tdc.codemp and vta.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "'"
  '  sql = sql & " where vta.codemp='" & gsCodEmp & "' and vta.pdoano='" & gsAnoAct & "' and mespvs in (" & Left(cmbEjercicio.Text, 2) & ") order by vta.fevdoc, vta.codtdc,vta.serdoc,vta.nrodoc "
    sql = sql & " where vta.codemp='" & gsCodEmp & "' and vta.pdoano='" & gsAnoAct & "' and mespvs in (" & Left(cmbEjercicio.Text, 2) & ") order by vta.codtdc,vta.serdoc,vta.nrodoc "
  '  sql = sql & " where vta.codemp='" & gsCodEmp & "' and vta.pdoano='" & gsAnoAct & "' and mespvs in (" & Left(cmbEjercicio.Text, 2) & ") order by vta.coddro,vta.nrocpb,vta.codtdc,vta.serdoc,vta.nrodoc "
   Case Else
    Exit Sub
  End Select
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sql
    .Open
  End With
  gpEncabezadoRptLibros frmMain.rptMain, rTitulo, udFecha, True, 1, porstMRp
  
  With frmMain.rptMain
    If porstMRp.RecordCount > 0 Then
      .ReportFileName = gsRutRpt & sReporte
    Else
      '.ReportFileName = gsRutRpt & "LOSINOPEX.rpt"
      .ReportFileName = gsRutRpt & sReporte
    End If
    .WindowShowExportBtn = IIf(1, True, False)
    .ParameterFields(1) = "Fecha;" & IIf(chkImpFecha.Value = Checked, 1, 0) & ";true"
    .ParameterFields(2) = "RepLegal;" & gsRepEmp & ";true"
    .ParameterFields(3) = "Contador;" & gsConEmp & ";true"
    .ParameterFields(4) = "Moneda;" & sMoneda & ";true"
    .MarginLeft = unMargenIzquierdo
    .WindowState = crptMaximized
    .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
    .Action = 1
  End With

End Sub
Sub LibroCaja(cuenta As String)
Dim sMoneda As String, sMonedae As String

Dim mesactual As String
  
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sMonedae = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_TXT, TPOMON_NAC_TXT)
  
      
  mesactual = gsMesAct
  
  gsMesAct = Left(cmbEjercicio.Text, 2)
  
  gpCamposSaldos
  
  With porstMRp
    
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.CodCta, a.CodDro, a.NroCpb, a.FehOpe, a.MesPvs, "
    .Source = .Source & "(CASE a.MesPvs WHEN '1' THEN 'ENERO' WHEN '2' THEN 'FEBRERO' WHEN '3' THEN 'MARZO' WHEN '4' THEN 'ABRIL' WHEN '5' THEN 'MAYO' WHEN '6' THEN 'JUNIO' WHEN '7' THEN 'JULIO' WHEN '8' THEN 'AGOSTO' WHEN '9' THEN 'SETIEMBRE' WHEN '10' THEN 'OCTUBRE' WHEN '11' THEN 'NOVIEMBRE' WHEN '12' THEN 'DICIEMBRE' END) AS Tmes, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    .Source = .Source & "a.CodAux, b.RazAux, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, d.DetDro, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cCargo, "
    .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cAbono, "
    If Left(cmbEjercicio.Text, 2) = "00" Then
        .Source = .Source & "(" & 0 & ") AS cAntCtaDeb, "
        .Source = .Source & "(" & 0 & ") AS cAntCtaHab, "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaCar, "
        '.Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
    Else
        .Source = .Source & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb, "
        .Source = .Source & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab, "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") ELSE 0 END) AS cAntCtaCar, "
        '.Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
    End If
    .Source = .Source & "FROM ((((COCpbDet a "
    '.Source = .Source & "LEFT JOIN cobancab x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and a.coddro=x.coddro and a.nrocpb=x.nroban "
    '.Source = .Source & "LEFT JOIN cobco ON x.codemp=cobco.codemp AND x.codbco=cobco.codbco "
    '.Source = .Source & "LEFT JOIN bnmediopago on x.codemp=bnmediopago.codemp AND x.tpodoc=bnmediopago.codmed "
    .Source = .Source & "LEFT JOIN bnmediopago on a.codemp=bnmediopago.codemp AND a.tpodoc=bnmediopago.codmed "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    .Source = .Source & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    If cuenta = "999" Then
        .Source = .Source & "AND left(a.CodCta,2) =10 "
    Else
        If cuenta = "101" Then
        .Source = .Source & "AND left(a.CodCta,3) in ('101','102','103')"
        Else
        .Source = .Source & "AND left(a.CodCta,3) in ('104','105','106','107')"
        End If
    End If
    .Source = .Source & "AND a.MesPvs ='" & Left(cmbEjercicio.Text, 2) & "' "
    '.Source = .Source & "ORDER BY a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
    '*****************************************************************************************************************
    .Source = .Source & " UNION ALL SELECT a.CodCta, '' as coddro, '' as NroCpb, FehOpe as FehOpe,'" & Left(cmbEjercicio.Text, 2) & "' as MesPvs, "
    .Source = .Source & "(CASE '" & Left(cmbEjercicio.Text, 2) & "' WHEN '1' THEN 'ENERO' WHEN '2' THEN 'FEBRERO' WHEN '3' THEN 'MARZO' WHEN '4' THEN 'ABRIL' WHEN '5' THEN 'MAYO' WHEN '6' THEN 'JUNIO' WHEN '7' THEN 'JULIO' WHEN '8' THEN 'AGOSTO' WHEN '9' THEN 'SETIEMBRE' WHEN '10' THEN 'OCTUBRE' WHEN '11' THEN 'NOVIEMBRE' WHEN '12' THEN 'DICIEMBRE' END) AS Tmes, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    .Source = .Source & "a.CodAux, b.RazAux, a.RefDoc, '' AS GloIte, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, '' as DetDro, "
    '.Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    '.Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    '.Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cCargo, "
    '.Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cAbono, "
    .Source = .Source & "0 AS cDebe, "
    .Source = .Source & "0 AS cHaber, "
    .Source = .Source & "0 AS cCargo, "
    .Source = .Source & "0 AS cAbono, "
    If Left(cmbEjercicio.Text, 2) = "00" Then
        .Source = .Source & "(" & 0 & ") AS cAntCtaDeb, "
        .Source = .Source & "(" & 0 & ") AS cAntCtaHab, "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaCar, "
        '.Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
    Else
        .Source = .Source & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb, "
        .Source = .Source & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab, "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") ELSE 0 END) AS cAntCtaCar, "
        '.Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        .Source = .Source & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
    End If
    .Source = .Source & "FROM ((((COCpbDet a "
    '.Source = .Source & "LEFT JOIN cobancab x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and a.coddro=x.coddro and a.nrocpb=x.nroban "
    '.Source = .Source & "LEFT JOIN cobco ON x.codemp=cobco.codemp AND x.codbco=cobco.codbco "
    '.Source = .Source & "LEFT JOIN bnmediopago on x.codemp=bnmediopago.codemp AND x.tpodoc=bnmediopago.codmed "
    .Source = .Source & "LEFT JOIN bnmediopago on a.codemp=bnmediopago.codemp AND a.tpodoc=bnmediopago.codmed "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    .Source = .Source & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
    .Source = .Source & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    If cuenta = "999" Then
        .Source = .Source & "AND left(a.CodCta,2) =10 "
    Else
        If cuenta = "101" Then
        .Source = .Source & "AND left(a.CodCta,3) in ('101','102','103')"
        Else
        .Source = .Source & "AND left(a.CodCta,3) in ('104','105','106','107')"
        End If
    End If
    .Source = .Source & "AND a.MesPvs <'" & Left(cmbEjercicio.Text, 2) & "' AND (COCtaAcu.AcuD" & Left(cmbEjercicio.Text, 2) & "_MN+COCtaAcu.AcuD" & Left(cmbEjercicio.Text, 2) & "_ME+COCtaAcu.AcuH" & Left(cmbEjercicio.Text, 2) & "_MN+COCtaAcu.AcuH" & Left(cmbEjercicio.Text, 2) & "_ME)=0 "
    '.Source = .Source & " ORDER BY a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
    .Source = .Source & " group by a.CodCta ORDER BY 5, 1, 2, 3 "
    .Open
  End With
  
  gpEncabezadoRptLibros frmMain.rptMain, IIf(cuenta = "101", "Formato 1.1: Libro Caja y Bancos detalle de los Movimientos del Efectivo", IIf(cuenta = "999", "FORMATO 3.2: LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 10 CAJA Y BANCOS", "Formato 1.2: Libro Caja y Bancos detalle de los Movimientos de la Cuenta Corriente")) & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & ")", udFecha, True, 1, porstMRp
      
  gsMesAct = mesactual
  
    With frmMain.rptMain
      
       '[Datos y parámetros del reporte.  'Cambiar.
      
      If porstMRp.RecordCount > 0 Then
      
      
      
      If cuenta = "999" Then
        .ReportFileName = gsRutRpt & "LOCaja10.rpt"
      Else
        .ReportFileName = gsRutRpt & "LOCaja.rpt"
      End If
      
      
      
      Else
      .ReportFileName = gsRutRpt & "LOSINOPE.rpt"
      End If
            
            
      'Parametros adicionales
      
      .ParameterFields(1) = "Equivalente;" & sEquivalente & ";true"
      .ParameterFields(2) = "Fecha;" & IIf(chkImpFecha.Value = Checked, 1, 0) & ";true"
      .ParameterFields(3) = "RepLegal;" & gsRepEmp & ";true"
      .ParameterFields(4) = "Contador;" & gsConEmp & ";true"

      '.WindowShowGroupTree = True
      
      .WindowState = crptMaximized
      .WindowShowExportBtn = True
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
      
    End With

End Sub
Private Function LibroCajaElectro(cuenta As String) As String
'Sub LibroCajaElectro(cuenta As String)
Dim sMoneda As String, sMonedae As String

Dim mesactual As String
  
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sMonedae = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_TXT, TPOMON_NAC_TXT)
  
      
  mesactual = gsMesAct
  
  gsMesAct = Left(cmbEjercicio.Text, 2)
  
  gpCamposSaldos
  Dim sSentencia1  As String
  sSentencia1 = ""
  '+With porstMRp
    
  '+  If .State = adStateOpen Then .Close
    sSentencia1 = "SELECT a.CodCta, a.CodDro, a.NroCpb, a.FehOpe, a.MesPvs, "
    sSentencia1 = sSentencia1 & "(CASE a.MesPvs WHEN '1' THEN 'ENERO' WHEN '2' THEN 'FEBRERO' WHEN '3' THEN 'MARZO' WHEN '4' THEN 'ABRIL' WHEN '5' THEN 'MAYO' WHEN '6' THEN 'JUNIO' WHEN '7' THEN 'JULIO' WHEN '8' THEN 'AGOSTO' WHEN '9' THEN 'SETIEMBRE' WHEN '10' THEN 'OCTUBRE' WHEN '11' THEN 'NOVIEMBRE' WHEN '12' THEN 'DICIEMBRE' END) AS Tmes, "
    sSentencia1 = sSentencia1 & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    sSentencia1 = sSentencia1 & "a.CodAux, b.RazAux, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, d.DetDro, "
    sSentencia1 = sSentencia1 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    sSentencia1 = sSentencia1 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cCargo, "
    sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cAbono, "
    If Left(cmbEjercicio.Text, 2) = "00" Then
        sSentencia1 = sSentencia1 & "(" & 0 & ") AS cAntCtaDeb, "
        sSentencia1 = sSentencia1 & "(" & 0 & ") AS cAntCtaHab, "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaCar, "
        'sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaAbo "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & ", bnmediopago.codmed,c.codbco,a.tpomon "
    Else
        sSentencia1 = sSentencia1 & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb, "
        sSentencia1 = sSentencia1 & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab, "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") ELSE 0 END) AS cAntCtaCar, "
        'sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") ELSE 0 END) AS cAntCtaAbo "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & ", bnmediopago.codmed,c.codbco,a.tpomon "
    End If
'ini 2015-04-27 corrige rpt
    sSentencia1 = sSentencia1 & ", bnmediopago.codmed,IFNULL(c.codbco,'') codbco ,a.tpomon "
    sSentencia1 = sSentencia1 & ",c.tpomon tpomoncta,IF(c.tpomon='N',IFNULL(f.ctactemn,''),IFNULL(f.ctacteme,'')) AS nrocta "
'fin 2015-04-27 corrige rpt

    '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & "FROM ((((COCpbDet a "
    sSentencia1 = sSentencia1 & "FROM (((((COCpbDet a "
    'sSentencia1 = sSentencia1 & "LEFT JOIN cobancab x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and a.coddro=x.coddro and a.nrocpb=x.nroban "
    'sSentencia1 = sSentencia1 & "LEFT JOIN cobco ON x.codemp=cobco.codemp AND x.codbco=cobco.codbco "
    'sSentencia1 = sSentencia1 & "LEFT JOIN bnmediopago on x.codemp=bnmediopago.codemp AND x.tpodoc=bnmediopago.codmed "
    sSentencia1 = sSentencia1 & "LEFT JOIN bnmediopago on a.codemp=bnmediopago.codemp AND a.tpodoc=bnmediopago.codmed "
    sSentencia1 = sSentencia1 & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    sSentencia1 = sSentencia1 & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    sSentencia1 = sSentencia1 & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    sSentencia1 = sSentencia1 & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
    sSentencia1 = sSentencia1 & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta "
'ini 2015-04-27 corrige rpt
    sSentencia1 = sSentencia1 & "LEFT JOIN CoBco f ON a.codemp=f.codemp AND c.codbco=f.codbco)"
'fin 2015-04-27 corrige rpt
    sSentencia1 = sSentencia1 & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia1 = sSentencia1 & "AND a.pdoano='" & gsAnoAct & "' "
    If cuenta = "999" Then
        sSentencia1 = sSentencia1 & "AND left(a.CodCta,2) =10 "
    Else
        If cuenta = "101" Then
        sSentencia1 = sSentencia1 & "AND left(a.CodCta,3) in ('101','102','103')"
        Else
        sSentencia1 = sSentencia1 & "AND left(a.CodCta,3) in ('104','105','106','107')"
        End If
    End If
    sSentencia1 = sSentencia1 & "AND a.MesPvs ='" & Left(cmbEjercicio.Text, 2) & "' "
    'sSentencia1 = sSentencia1 & "ORDER BY a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
    '*****************************************************************************************************************
    sSentencia1 = sSentencia1 & " UNION ALL SELECT a.CodCta, '' as coddro, '' as NroCpb, FehOpe as FehOpe,'" & Left(cmbEjercicio.Text, 2) & "' as MesPvs, "
    sSentencia1 = sSentencia1 & "(CASE '" & Left(cmbEjercicio.Text, 2) & "' WHEN '1' THEN 'ENERO' WHEN '2' THEN 'FEBRERO' WHEN '3' THEN 'MARZO' WHEN '4' THEN 'ABRIL' WHEN '5' THEN 'MAYO' WHEN '6' THEN 'JUNIO' WHEN '7' THEN 'JULIO' WHEN '8' THEN 'AGOSTO' WHEN '9' THEN 'SETIEMBRE' WHEN '10' THEN 'OCTUBRE' WHEN '11' THEN 'NOVIEMBRE' WHEN '12' THEN 'DICIEMBRE' END) AS Tmes, "
    sSentencia1 = sSentencia1 & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    sSentencia1 = sSentencia1 & "a.CodAux, b.RazAux, a.RefDoc, '' AS GloIte, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, '' as DetDro, "
    'sSentencia1 = sSentencia1 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    'sSentencia1 = sSentencia1 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    'sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cCargo, "
    'sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN (CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMonedae & " ELSE 0 END) ELSE 0 END) AS cAbono, "
    sSentencia1 = sSentencia1 & "0 AS cDebe, "
    sSentencia1 = sSentencia1 & "0 AS cHaber, "
    sSentencia1 = sSentencia1 & "0 AS cCargo, "
    sSentencia1 = sSentencia1 & "0 AS cAbono, "
    If Left(cmbEjercicio.Text, 2) = "00" Then
        sSentencia1 = sSentencia1 & "(" & 0 & ") AS cAntCtaDeb, "
        sSentencia1 = sSentencia1 & "(" & 0 & ") AS cAntCtaHab, "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaCar, "
        'sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & 0 & ") ELSE 0 END) AS cAntCtaAbo "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & ", bnmediopago.codmed,c.codbco,a.tpomon "
    Else
        sSentencia1 = sSentencia1 & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb, "
        sSentencia1 = sSentencia1 & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab, "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") ELSE 0 END) AS cAntCtaCar, "
        'sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 4, 3)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.abvmed,detbco,(CASE x.tpomon WHEN 'N' THEN cobco.ctactemn else cobco.ctacteme end) as cuenta  "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") ELSE 0 END) AS cAntCtaAbo, bnmediopago.codmed,c.codbco,a.tpomon "
        sSentencia1 = sSentencia1 & "(CASE c.TpoMon WHEN '" & Right(sMonedae, 1) & "' THEN " & "(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") ELSE 0 END) AS cAntCtaAbo "
        '2015-04-27 corrige rpt sSentencia1 = sSentencia1 & ", bnmediopago.codmed,c.codbco,a.tpomon "
    End If
'ini 2015-04-27 corrige rpt
    sSentencia1 = sSentencia1 & ", bnmediopago.codmed,IFNULL(c.codbco,'') codbco,a.tpomon "
    sSentencia1 = sSentencia1 & ",c.tpomon tpomoncta,IF(c.tpomon='N',IFNULL(f.ctactemn,''),IFNULL(f.ctacteme,'')) AS nrocta "
'fin 2015-04-27 corrige rpt
    '2015-04-27 corrige rpt  sSentencia1 = sSentencia1 & "FROM ((((COCpbDet a "
    sSentencia1 = sSentencia1 & "FROM (((((COCpbDet a "
    'sSentencia1 = sSentencia1 & "LEFT JOIN cobancab x on a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and a.coddro=x.coddro and a.nrocpb=x.nroban "
    'sSentencia1 = sSentencia1 & "LEFT JOIN cobco ON x.codemp=cobco.codemp AND x.codbco=cobco.codbco "
    'sSentencia1 = sSentencia1 & "LEFT JOIN bnmediopago on x.codemp=bnmediopago.codemp AND x.tpodoc=bnmediopago.codmed "
    sSentencia1 = sSentencia1 & "LEFT JOIN bnmediopago on a.codemp=bnmediopago.codemp AND a.tpodoc=bnmediopago.codmed "
    sSentencia1 = sSentencia1 & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    sSentencia1 = sSentencia1 & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    sSentencia1 = sSentencia1 & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    sSentencia1 = sSentencia1 & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
    sSentencia1 = sSentencia1 & "LEFT JOIN COCtaAcu ON a.codemp=COCtaAcu.codemp AND a.pdoano=COCtaAcu.pdoano AND a.CodCta=COCtaAcu.CodCta "
'ini 2015-04-27 corrige rpt
    sSentencia1 = sSentencia1 & "LEFT JOIN CoBco f ON a.codemp=f.codemp AND c.codbco=f.codbco)"
'fin 2015-04-27 corrige rpt
    sSentencia1 = sSentencia1 & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia1 = sSentencia1 & "AND a.pdoano='" & gsAnoAct & "' "
    If cuenta = "999" Then
        sSentencia1 = sSentencia1 & "AND left(a.CodCta,2) =10 "
    Else
        If cuenta = "101" Then
        sSentencia1 = sSentencia1 & "AND left(a.CodCta,3) in ('101','102','103')"
        Else
        sSentencia1 = sSentencia1 & "AND left(a.CodCta,3) in ('104','105','106','107')"
        End If
    End If
    sSentencia1 = sSentencia1 & "AND a.MesPvs <'" & Left(cmbEjercicio.Text, 2) & "' AND (COCtaAcu.AcuD" & Left(cmbEjercicio.Text, 2) & "_MN+COCtaAcu.AcuD" & Left(cmbEjercicio.Text, 2) & "_ME+COCtaAcu.AcuH" & Left(cmbEjercicio.Text, 2) & "_MN+COCtaAcu.AcuH" & Left(cmbEjercicio.Text, 2) & "_ME)=0 "
    'sSentencia1 = sSentencia1 & " ORDER BY a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
    sSentencia1 = sSentencia1 & " group by a.CodCta ORDER BY 5, 1, 2, 3 "
    '+.Open
  '+End With
  
  '+gpEncabezadoRptLibros frmMain.rptMain, IIf(cuenta = "101", "Formato 1.1: Libro Caja y Bancos detalle de los Movimientos del Efectivo", IIf(cuenta = "999", "FORMATO 3.2: LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 10 CAJA Y BANCOS", "Formato 1.2: Libro Caja y Bancos detalle de los Movimientos de la Cuenta Corriente")) & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & ")", udFecha, True, 1, porstMRp
      
  gsMesAct = mesactual
  
    '+With frmMain.rptMain
      
       '[Datos y parámetros del reporte.  'Cambiar.
      
      '+If porstMRp.RecordCount > 0 Then
      
      
'+
''      If cuenta = "999" Then
''        .ReportFileName = gsRutRpt & "LOCaja10.rpt"
''      Else
''        .ReportFileName = gsRutRpt & "LOCaja.rpt"
''      End If
      
      
''      Else
''      .ReportFileName = gsRutRpt & "LOSINOPE.rpt"
''      End If
'+
            
'+
''      'Parametros adicionales
''
''      .ParameterFields(1) = "Equivalente;" & sEquivalente & ";true"
''      .ParameterFields(2) = "Fecha;" & IIf(chkImpFecha.Value = Checked, 1, 0) & ";true"
''      .ParameterFields(3) = "RepLegal;" & gsRepEmp & ";true"
''      .ParameterFields(4) = "Contador;" & gsConEmp & ";true"
''
''      '.WindowShowGroupTree = True
''
''      .WindowState = crptMaximized
''      .WindowShowExportBtn = True
''      .MarginLeft = unMargenIzquierdo
''      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
''      .Action = 1
''
''    End With
'+
'End Sub
    LibroCajaElectro = sSentencia1
End Function

Sub LibroMayor()
  Dim nContador As Integer, sMoneda As String
  Dim s_MesIni As String, s_MesFin As String
  Dim s_SalAno As String, s_SalMes As String
  Dim s_Sentencia As String, s_Sql As String
  Dim l_CreateTB As Boolean, n_Index As Integer
  Dim s_Catalogo As String
  Dim sSalAntDeb As String, sSalAntHab As String
  
  s_MesIni = Left(cmbEjercicio.Text, 2)
  s_MesFin = Left(cmbEjercicio.Text, 2)
  ' Valido el rango de periodos
  'sMoneda = TPOMON_NAC_TXT
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "MN", "ME")
  ' Cadena de saldo anterior
  With porstMRp
    If .State = adStateOpen Then .Close
    s_Catalogo = "CoCtaAcu"
    s_Sentencia = "SELECT a.MesPvs AS MesPvs, a.CodCta AS CodCta, a.CodDro AS CodDro, a.NroCpb AS NroCpb, a.NroIte AS NroIte, a.FehOpe, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    s_Sentencia = s_Sentencia & "a.CodAux, b.RazAux, a.RefDoc, a.tpodoc, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta , " & Choose(gsIdioma, "d.DetDro", "d.DetDrox") & " AS DetDro, e.AbvTDc, "
    If s_MesIni <> "00" Then
      sSalAntDeb = "ROUND(("
      sSalAntHab = "ROUND(("
      s_SalMes = s_MesIni
      For nContador = 0 To (Val(s_SalMes) - 1)
        sSalAntDeb = sSalAntDeb & "acu.AcuD" & Format(nContador, "00") & "_" & sMoneda & IIf(nContador = (Val(s_SalMes) - 1), ")", "+")
        sSalAntHab = sSalAntHab & "acu.AcuH" & Format(nContador, "00") & "_" & sMoneda & IIf(nContador = (Val(s_SalMes) - 1), ")", "+")
      Next nContador
      sSalAntDeb = sSalAntDeb & ", 2)"
      sSalAntHab = sSalAntHab & ", 2)"
      s_Sentencia = s_Sentencia & sSalAntDeb & " AS cAntCtaDeb, "
      s_Sentencia = s_Sentencia & sSalAntHab & " AS cAntCtaHab "
    Else
      s_Sentencia = s_Sentencia & "0 AS cAntCtaDeb, 0 AS cAntCtaHab "
    End If
    s_Sentencia = s_Sentencia & "FROM ((((COCpbDet a "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    s_Sentencia = s_Sentencia & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & s_Catalogo & " acu ON a.codemp=acu.codemp AND a.pdoano=acu.pdoano AND a.CodCta=acu.CodCta "
    s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND a.pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND a.MesPvs>='" & s_MesIni & "' AND a.MesPvs<='" & s_MesFin & "' "
    s_Sentencia = s_Sentencia & "AND NOT(a.tpognr IN('" & TPOGNR_DST & "', '" & TPOGNR_DCA & "') AND a.imp" & sMoneda & "=0.00) "
    If s_MesIni <> "00" Then
      s_Catalogo = "CoCtaAcu"
      s_Sentencia = s_Sentencia & "UNION "
      s_Sentencia = s_Sentencia & "SELECT '00' AS MesPvs, c.CodCta AS CodCta, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, 0, 0, "
      s_Sentencia = s_Sentencia & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta , Null, Null, "
      s_Sentencia = s_Sentencia & sSalAntDeb & " AS cAntCtaDeb, "
      s_Sentencia = s_Sentencia & sSalAntHab & " AS cAntCtaHab "
      s_Sentencia = s_Sentencia & "FROM (COCta c "
      s_Sentencia = s_Sentencia & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.CodCta=a.CodCta) "
      s_Sentencia = s_Sentencia & "LEFT JOIN " & s_Catalogo & " acu ON c.codemp=acu.codemp AND c.pdoano=acu.pdoano AND c.CodCta=acu.CodCta "
      s_Sentencia = s_Sentencia & "WHERE c.codemp='" & gsCodEmp & "' "
      s_Sentencia = s_Sentencia & "AND c.pdoano='" & gsAnoAct & "' "
      s_Sentencia = s_Sentencia & "AND c.TpoCta='" & TPOCTA_TRA & "' "
      If ps_Plataforma = pSrvMySql Then
        s_Sentencia = s_Sentencia & "HAVING (ROUND(cAntCtaDeb, 2)<>0.00 OR ROUND(cAntCtaHab, 2)<>0.00) "
      Else
        s_Sentencia = s_Sentencia & "AND (ROUND(" & sSalAntDeb & ", 2)<>0.00 OR ROUND(" & sSalAntHab & ", 2)<>0.00) "
      End If
    End If
    s_Sentencia = s_Sentencia & "ORDER BY CodCta, MesPvs, CodDro, NroCpb, NroIte"
    .Source = s_Sentencia
    .Open
  End With

  s_Sentencia = ""
  'gpEncabezadoRpt frmMain.rptMain, "Formato 6.1: Libro Mayor" & " (" & TPOMON_NAC_TXT_1 & ")", udFecha, True, 1, porstMRp
  gpEncabezadoRpt frmMain.rptMain, "Formato 6.1: Libro Mayor" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & ")", udFecha, True, 1, porstMRp
 
  
  With frmMain.rptMain
    '2015-05-28 cambio gsMesAct a s_MesIni  .Formulas(5) = "mPeriodo='" & s_Sentencia & " " & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
    .Formulas(5) = "mPeriodo='" & s_Sentencia & " " & gfMesLet("01" & s_MesIni & gsAnoAct, 0, "", 1, " ", 1) & "'"
    .ParameterFields(1) = "Fecha;" & IIf(chkImpFecha.Value = Checked, 1, 0) & ";true"
    
    .ReportFileName = gsRutRpt & "LOMayor.rpt"
    .WindowShowExportBtn = True
    .WindowState = crptMaximized
    .MarginLeft = unMargenIzquierdo
    .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
    .Action = 1
  End With
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpSaldosIni", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpSaldosIni') DROP TABLE #tmpSaldosIni")
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpSaldosApe", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpSaldosApe') DROP TABLE #tmpSaldosApe")

End Sub
Private Sub ppArchivoElectronico(ByVal sArchivo As String, ByVal sNombreArchivo As String, ByVal sSentencia As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  '2016-02-02.03 correccion ple Dim psRegistro As String, s_Caracter As String, s_Expresion As String
  Dim psRegistro As String, s_Expresion As String
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nRegistroAux As Long, nRegistroDeta As Long, nTamano As Integer
  Dim sAuxiliar As String, s_OldMessage As String
  Dim nSumatoriaTotal As Double
  
   Dim n_SdoDeb As Double, n_SdoHab As Double 'teo 2015-12-15 falta definir
 
  ' selecciono informacion de proceso
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
    nRegistros = .RecordCount
  End With
  
    Select Case pnOpcion
    Case 19
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
    End Select

  
  ' Creo objeto de archivo
  If nRegistros > 0 Then
    s_Expresion = Left(sNombreArchivo, 30) & "1" & Mid(sNombreArchivo, 32)
    sArchivo = Replace(sArchivo, sNombreArchivo, s_Expresion)
'ini 2016-02-02.26 correccion ple archivo vacio=0
  Else
    s_Expresion = Left(sNombreArchivo, 30) & "0" & Mid(sNombreArchivo, 32)
    sArchivo = Replace(sArchivo, sNombreArchivo, s_Expresion)
'fin 2016-02-02.26 correccion ple archivo vacio=0
  End If
  Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
  Set potxtFileExp = pofsoFileExp.CreateTextFile(sArchivo, True)
  s_Caracter = "|"
  
  ' detalle de archivo
  Dim xxMesPvs As String
  xxMesPvs = Left(cmbEjercicio.Text, 2)
  'ini 2014-07-31 numero correlativo
  '2015-04-07 error desbordamiento Dim xNroCorr As Integer
  Dim xNroCorr As Long
  xNroCorr = 1
  'fin 2014-07-31 numero correlativo
  If Not (porstMRp.BOF And porstMRp.EOF) Then
    nRegistro = 0
    While Not porstMRp.EOF
      psRegistro = ""
      Select Case pnOpcion
       Case 5       ' 3.2 libro caja
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!codbco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!NroCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!TpoMonCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        '2016-02-05 error option explic nImporte_mn = cDebe
        nImporte_mn = porstMRp!cDebe
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '2016-02-05 error option explic nImporte_mn = porstMRp!cHaber
        nImporte_mn = porstMRp!cHaber
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
       '2015-11-12 PLE rpt fmt electro archivo  Case 6       ' 3.2 libro caja
       '2015-11-12/18 PLE rpt fmt electro archivo  Case 6, 7      ' 3.2 libro caja,
'case=6y7 3.3 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 12 CUENTAS POR COBRAR COMERCIALES – TERCEROS Y 13
'case=8 3.4 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS
'case=9y10 3.5 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS - TERCEROS O CUENTA 17
'case=11 3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
'case=12 3.11 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 41 REMUNERACIONES Y PARTICIPACIONES POR PAGAR (PCGE) (2)
'case=13y14 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=15y16 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=18 3.15  LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 37 ACTIVO DIFERIDO Y DE LA CUENTA 49 PASIVO DIFERIDO (PCGE)   (2)

'archivo agrupado y ordenado por RUC
       Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
             
'ini 2015-11-12/18 PLE rpt fmt electro archivo
'        s_Expresion = "NoExiste"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2015-11-12/18 PLE rpt fmt electro archivo
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 12
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
'fin 2015-12-07 libros inven y balance correcc
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = porstMRp!rucaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 13, 14, 15, 16
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
'fin 2015-12-07 libros inven y balance correcc
        
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 12
        s_Expresion = porstMRp!codaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
        
'fin 2015-12-07 libros inven y balance correcc
        s_Expresion = porstMRp!razAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2015-12-07 libros inven y balance correcc
''        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
''        psRegistro = psRegistro & s_Expresion & s_Caracter
        Select Case pnOpcion
        Case 12
        Case Else
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
 'fin 2015-12-07 libros inven y balance correcc
       
 'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 15, 16
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
'fin 2015-12-07 libros inven y balance correcc
      
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'ini 2015-12-04 libros inven y balance correcc
        s_Expresion = porstMRp!codtdc
       'psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & "-" & porstMRp!serdoc
       'psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & "-" & porstMRp!nrodoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
       Case 18
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2015-11-12/18 PLE rpt fmt electro archivo
'        s_Expresion = "NoExiste"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
'fin 2015-11-12/18 PLE rpt fmt electro archivo
'ini 2015-12-04 libros inven y balance correcc
        s_Expresion = porstMRp!codtdc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & porstMRp!serdoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & porstMRp!nrodoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!GloIte
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'adiciones
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        'deducciones
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
        
'ini 2015-11-12/18 PLE rpt fmt electro archivo


'fin 2015-12-04 libros inven y balance correcc
        
'ini 2015-11-12/13 PLE rpt fmt electro archivo
        
'''''ini 2015-12-07 libros inven y balance correcc
''''       Case 11   '3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
''''        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        ' 2: numero correlativo o codigo unico
''''        s_Expresion = "0000-000000-000000"
''''        If Not IsNull(porstMRp!coddro) Then
''''          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
''''        End If
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
''''         ' 3: segun nueva estructura numero correlativo de asiento contable
''''         s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
''''         psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
''''    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
''''
''''        s_Expresion = porstMRp!TpoDci
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!rucaux
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!razAux
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!codtdc
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!serdoc
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!nrodoc
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        nImporte_mn = CDec(porstMRp!DebeSol)
''''        nImporte_me = CDec(porstMRp!HaberSol)
''''        n_Importe = Round(nImporte_mn - nImporte_me, 2)
''''        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''         s_Expresion = "1"
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
'''''fin 2015-12-07 libros inven y balance correcc
        
'ini 2015-12-10 PLE rpt fmt electro archivo
       Case 19  '3.17 LIBRO DE INVENTARIOS Y BALANCES - BALANCE DE COMPROBACIÓN (3)
       'Periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Código de la cuenta
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Saldos iniciales Debe
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cApeD), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Saldos iniciales Haber
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cApeH), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Movimientos del ejercicio o periodo - Debe
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cSumaD), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Movimientos del ejercicio o periodo - Haber
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cSumaH), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'teo 2015-12-15 falta definir
        psRegistro = psRegistro & "falta suma del mayo debe" & s_Caracter
        psRegistro = psRegistro & "falta suma del mayo haber" & s_Caracter
        'teo 2015-12-15 falta definir
        
       'Saldos al 31 de Diciembre - Deudor (saldos finales deudor)
'ini 2016-02-02.03 correccion ple
'        nImporte_mn = CDec(porstMRp!cApeD + cSumaD)
'        nImporte_me = CDec(porstMRp!cApeH + cSumaH)
'error option explicit
        nImporte_mn = CDec(porstMRp!cApeD + porstMRp!cSumaD)
        nImporte_me = CDec(porstMRp!cApeH + porstMRp!cSumaH)
'fin 2016-02-02.03 correccion ple
        n_Importe = Round(IIf(nImporte_mn > nImporte_me, nImporte_mn - nImporte_me, 0), 2)
        n_SdoDeb = n_Importe 'teo 2015-12-15 falta definir
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Saldos al 31 de Diciembre - Acreedor (saldos finales deudor)
'ini 2016-02-02.03 correccion ple
'        nImporte_mn = CDec(porstMRp!cApeD + cSumaD)
'        nImporte_me = CDec(porstMRp!cApeH + cSumaH)
'error option explict
        nImporte_mn = CDec(porstMRp!cApeD + porstMRp!cSumaD)
        nImporte_me = CDec(porstMRp!cApeH + porstMRp!cSumaH)
'fin 2016-02-02.03 correccion ple
        n_Importe = Round(IIf(nImporte_mn > nImporte_me, 0, nImporte_mn - nImporte_me), 2)
        n_SdoHab = n_Importe 'teo 2015-12-15 falta definir
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'teo 2015-12-15 falta definir
        psRegistro = psRegistro & "falta Transferencias y Cancelaciones - Debe" & s_Caracter
        psRegistro = psRegistro & "falta Transferencias y Cancelaciones - Haber" & s_Caracter
        psRegistro = psRegistro & "falta Cuentas de Balance - Activo" & s_Caracter
        psRegistro = psRegistro & "falta Cuentas de Balance - Pasivo" & s_Caracter
        'teo 2015-12-15 falta definir
        
        'Resultado por Naturaleza - Pérdidas / Sdo. Finales del estado de perdidas y ganancia por funcion peridida
        n_Importe = Round(IIf(n_SdoDeb > 0 And (porstMRp!TpoSdo = "F" Or porstMRp!TpoSdo = "A"), n_SdoDeb, 0), 2)
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'Resultado por Naturaleza - Ganancias / Sdo. Finales del estado de perdidas y ganancia por funcion ganancia
        n_Importe = Round(IIf(n_SdoHab > 0 And (porstMRp!TpoSdo = "F" Or porstMRp!TpoSdo = "A"), n_SdoHab, 0), 2)
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
        'teo 2015-12-15 falta definir
        psRegistro = psRegistro & "falta Adiciones" & s_Caracter
        psRegistro = psRegistro & "falta Deducciones" & s_Caracter
        'teo 2015-12-15 falta definir
       
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
       
'fin 2015-12-10 PLE rpt fmt electro archivo
       Case 23: psRegistro = fAE_05_1_lib_diario(xNroCorr)        ' libro diario
       Case 24: psRegistro = fAE_05_2_lib_diario_simplificado(xNroCorr)        'libro diario  simplificado
       Case 27: psRegistro = fAE_06_1_lib_mayor(xNroCorr)        ' libro mayor
       Case 29, 30: psRegistro = fAE_08_1_reg_cpr(xNroCorr)   ' registro compras
       Case 31: psRegistro = fAE_14_1_reg_vta(xNroCorr)    ' resgistro ventas
'ini 2014-05-30 adicion 5.3 plan ctas
       Case 32: psRegistro = fAE_05_3_lib_diario_deta_plan_cta(xNroCorr)    ' plan de cuentas
      End Select
      potxtFileExp.WriteLine psRegistro
      nRegistro = nRegistro + 1
      porstMRp.MoveNext
      xNroCorr = xNroCorr + 1 '2014-07-31 numero correlativo
    Wend
    ' Cierro objeto y saco de memoria
    potxtFileExp.Close
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  GoTo Finalizar
  
error:
Finalizar:
  ' Reinicializo los mensajes
  ' Coloco el puntero en normal

End Sub
'Private Function fAE_08_1_reg_cpr_sql(sMoneda As String) As String
Private Function fAE_08_1_reg_cpr_sql(sMoneda As String, xFiltro As String) As String
'xFiltro= para 8.1 tipo doc <> a 91,97,98,
'xFiltro= para 8.2 tipo doc = a 00,91,97,98,
    Dim sSentencia As String
    
    sSentencia = "SELECT concat(com.coddro,com.nrocpb) as nrocpb, date_format(com.feedoc,'%d/%m/%Y') as feedoc, date_format(com.fevdoc,'%d/%m/%Y') as fevdoc, com.codtdc as codtdc, "
    sSentencia = sSentencia & "IFNULL(case com.codtdc when '50' then com.codaduana when '52' then com.codaduana when '53' then com.codaduana else com.serdoc end, '-') as serdoc, "
    sSentencia = sSentencia & "IFNULL(com.annodua, '0') as anno, IFNULL(CASE com.codtdc when '50' then com.nrodua when '52' then com.nrodua when '53' then com.nrodua else com.nrodoc END, '') as nrodoc, "
    sSentencia = sSentencia & "right(aux.tpodci,1) as tpodci, com.codaux as codaux, aux.razaux as razaux, "
    sSentencia = sSentencia & "com.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp1, "
    sSentencia = sSentencia & "com.impigv_ogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp2, "
    sSentencia = sSentencia & "com.impogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp3, "
    sSentencia = sSentencia & "com.impigv_ogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp4, "
    sSentencia = sSentencia & "com.impong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp5, "
    sSentencia = sSentencia & "com.impigv_ong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp6, "
    sSentencia = sSentencia & "com.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp7, "
    sSentencia = sSentencia & "com.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp8, "
    sSentencia = sSentencia & "com.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp9, "
    sSentencia = sSentencia & "com.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp10, "
    sSentencia = sSentencia & "CASE WHEN com.codtdc_ref='91' THEN Concat(com.serdoc_ref, '-', com.nrodoc_ref) ELSE Null END as v1, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "com.nrocdt, com.fehcdt, com.imptcb, com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
    sSentencia = sSentencia & "com.nrocdt, com.fehcdt,"
    sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN com.imptcb ELSE 0 END imptcb,"
    sSentencia = sSentencia & "com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
'fin 2015-09-16 t.cam solo ME
    'ini 2014-05-25
    sSentencia = sSentencia & ",codaduana,annodua,nrodua "
    'fin 2014-05-25
'ini 2016-02-02.05 correccion ple
    sSentencia = sSentencia & ",tpomon,TRIM(" & fIsNull() & "TpoBns,'')) TpoBns ,TRIM(" & fIsNull() & "CodMon,'')) CodMon "
    sSentencia = sSentencia & ",TRIM(" & fIsNull() & "PaisNoDomi,'')) PaisNoDomi "
    sSentencia = sSentencia & ",TRIM(" & fIsNull() & "PaisBenefi,'')) PaisBenefi,TRIM(" & fIsNull() & "CnveDobImpo,'')) CnveDobImpo "
    sSentencia = sSentencia & ",TRIM(" & fIsNull() & "TpoRta,'')) TpoRta,TasaReten ,TRIM(" & fIsNull() & "nrodociden,'')) nrodociden"
    '2016-02-02.09 sSentencia = sSentencia & ",aux.diraux,aux.rucaux "
    sSentencia = sSentencia & "," & fIsNull() & "aux.diraux,'') diraux," & fIsNull() & "aux.rucaux,'') rucaux "
'fin 2016-02-02.05 correccion ple
    sSentencia = sSentencia & ", date_format(com.fehope,'%d/%m/%Y') as fehope " '215-08-10 adicionado por teo
    sSentencia = sSentencia & "FROM cocprdoc com "
    sSentencia = sSentencia & "INNER JOIN tgaux aux ON com.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc ON com.codemp=tdc.codemp and com.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE com.codemp='" & gsCodEmp & "' AND com.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    If xFiltro = 1 Then
        sSentencia = sSentencia & "AND NOT com.codtdc IN('91','97','98') "
    Else
        sSentencia = sSentencia & "AND com.codtdc IN('00','91','97','98') "
    End If
    sSentencia = sSentencia & "ORDER BY 1"
    fAE_08_1_reg_cpr_sql = sSentencia
End Function
Private Function fAE_08_2_reg_cpr_sql(sMoneda As String, xFiltro As String) As String
'xFiltro= para 8.1 tipo doc <> a 91,97,98,
'xFiltro= para 8.2 tipo doc = a 00,91,97,98,
    Dim sSentencia As String
    
    sSentencia = "SELECT concat(com.coddro,com.nrocpb) as nrocpb, date_format(com.feedoc,'%d/%m/%Y') as feedoc, date_format(com.fevdoc,'%d/%m/%Y') as fevdoc, com.codtdc as codtdc, "
    sSentencia = sSentencia & "IFNULL(case com.codtdc when '50' then com.codaduana when '52' then com.codaduana when '53' then com.codaduana else com.serdoc end, '-') as serdoc, "
    sSentencia = sSentencia & "IFNULL(com.annodua, '0') as anno, IFNULL(CASE com.codtdc when '50' then com.nrodua when '52' then com.nrodua when '53' then com.nrodua else com.nrodoc END, '') as nrodoc, "
    'sSentencia = sSentencia & "right(aux.tpodci,1) as tpodci, com.codaux as codaux, aux.razaux as razaux, "
    sSentencia = sSentencia & "right(aux.tpodci,1) as tpodci,  aux.razaux as razaux, "
    sSentencia = sSentencia & "com.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp1, "
    sSentencia = sSentencia & "com.impigv_ogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp2, "
    sSentencia = sSentencia & "com.impogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp3, "
    sSentencia = sSentencia & "com.impigv_ogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp4, "
    sSentencia = sSentencia & "com.impong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp5, "
    sSentencia = sSentencia & "com.impigv_ong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp6, "
    sSentencia = sSentencia & "com.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp7, "
    sSentencia = sSentencia & "com.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp8, "
    sSentencia = sSentencia & "com.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp9, "
    sSentencia = sSentencia & "com.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp10, "
    sSentencia = sSentencia & "CASE WHEN com.codtdc_ref='91' THEN Concat(com.serdoc_ref, '-', com.nrodoc_ref) ELSE Null END as v1, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "com.nrocdt, com.fehcdt, com.imptcb, com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
    sSentencia = sSentencia & "com.nrocdt, com.fehcdt,"
    sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN com.imptcb ELSE 0 END imptcb,"
    sSentencia = sSentencia & "com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
'fin 2015-09-16 t.cam solo ME com.codaux as codaux,
    'ini 2014-05-25
    sSentencia = sSentencia & ",codaduana,annodua,nrodua "
    'fin 2014-05-25
'ini 2016-02-02.05 correccion ple
    sSentencia = sSentencia & ",tpomon,TRIM(" & fIsNull() & "TpoBns,'')) TpoBns ,TRIM(" & fIsNull() & "CodMon,'')) CodMon "
    sSentencia = sSentencia & ",TRIM(" & fIsNull() & "PaisNoDomi,'')) PaisNoDomi "
    sSentencia = sSentencia & ",TRIM(" & fIsNull() & "PaisBenefi,'')) PaisBenefi,TRIM(" & fIsNull() & "CnveDobImpo,'')) CnveDobImpo "
    sSentencia = sSentencia & ",TRIM(" & fIsNull() & "TpoRta,'')) TpoRta,TasaReten ,TRIM(" & fIsNull() & "nrodociden,'')) nrodociden"
    '2016-02-02.09 sSentencia = sSentencia & ",aux.diraux,aux.rucaux "
    sSentencia = sSentencia & "," & fIsNull() & "aux.diraux,'') diraux," & fIsNull() & "aux.rucaux,'') rucaux "
'fin 2016-02-02.05 correccion ple
    sSentencia = sSentencia & ", date_format(com.fehope,'%d/%m/%Y') as fehope " '215-08-10 adicionado por teo
    sSentencia = sSentencia & "FROM cocprdoc com "
    sSentencia = sSentencia & "INNER JOIN tgaux aux ON com.nrodociden=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc ON com.codemp=tdc.codemp and com.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE com.codemp='" & gsCodEmp & "' AND com.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    If xFiltro = 1 Then
        sSentencia = sSentencia & "AND NOT com.codtdc_ref IN('91','97','98') "
    Else
        sSentencia = sSentencia & "AND com.codtdc_ref IN('00','91','97','98') "
    End If
    sSentencia = sSentencia & "ORDER BY 1"
    fAE_08_2_reg_cpr_sql = sSentencia
End Function

Private Function fAE_08_2_reg_cpr(xNroCorr As Long) As String
  Dim s_Expresion As String
  Dim psRegistro As String
        '+ 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 3: Nueva version Contribuyentes del Régimen General: Número
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 4: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 5: Tipo de Comprobante de Pago o Documento del sujeto no domiciliado
        s_Expresion = porstMRp!codtdc_ref
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 6: Serie del comprobante de pago o documento.
        s_Expresion = porstMRp!serdoc_ref
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 7: Número del comprobante de pago o documento.
        s_Expresion = porstMRp!nrodoc_ref
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 8: Valor de las adquisiciones
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        '+ 22 - 9: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp9), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 23 - 10 : importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp1 + porstMRp!imp3 + porstMRp!imp5 + porstMRp!imp7), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter

        '+ 6 - 11: tipo comprobante de pago
        s_Expresion = Format(IIf(porstMRp!codtdc = "91", "", porstMRp!codtdc), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 7 - 12: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "", porstMRp!serdoc)
        's_Expresion = IIf(porstMRp!codtdc = "05", Right(porstMRp!serdoc, 1), porstMRp!serdoc) '2014-07-16
        s_Expresion = IIf(porstMRp!codtdc = "91", "", porstMRp!serdoc) '2014-07-16
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 8 - 13: año emision DUA
        s_Expresion = IIf(IsNull(porstMRp!anno), "0", porstMRp!anno)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 9 - 14: numero comprobante de pago
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
        IIf(porstMRp!codtdc = "91", "", porstMRp!nrodoc))
        'IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
                
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 15: Monto de retención del IGV
        's_Expresion = ""
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 16: Código  de la Moneda (Tabla 4)
        s_Expresion = porstMRp!codmon
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 17: Tipo de cambio
        'por teo 2016-02-10 s_Expresion = ""
        s_Expresion = Replace(FormatNumber(CDec(IIf(IsNull(porstMRp!ImpTCb), 0, porstMRp!ImpTCb)), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 18: Pais de la residencia del sujeto no domiciliado
        s_Expresion = porstMRp!PaisNoDomi
        psRegistro = psRegistro & s_Expresion & s_Caracter

        '+ 13 - 19: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 20: Domicilio en el extranjero del sujeto no domiciliado
        s_Expresion = porstMRp!DirAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 21: Número de identificación del sujeto no domiciliado
        s_Expresion = porstMRp!nrodociden
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 22: Número de identificación fiscal del beneficiario efectivo de los pagos
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 23: Apellidos y nombres, denominación o razón social  del beneficiario efectivo de los pagos
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 24: Pais de la residencia del beneficiario efectivo de los pagos
        s_Expresion = porstMRp!PaisBenefi
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 25: Vínculo entre el contribuyente y el residente en el extranjero
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 26: Renta Bruta
        s_Expresion = ""
        's_Expresion = Replace(FormatNumber(CDec(porstMRp!imp1 + porstMRp!imp2 + porstMRp!imp3 + porstMRp!imp4 + porstMRp!imp5 + porstMRp!imp6 + porstMRp!imp7), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 27: Deducción / Costo de Enajenación de bienes de capital
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 28: Renta Neta
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 29: Tasa de retención
        s_Expresion = porstMRp!TasaReten
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
         'nuevo 30: Impuesto retenido / dice:cacular 29 * 26 çrenta bruta), paro 29 es cero
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
        'nuevo 31: Convenios para evitar la doble imposición
        s_Expresion = porstMRp!CnveDobImpo
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
         'nuevo 32: Exoneración aplicada
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        'nuevo 33: Tipo de Renta
        s_Expresion = porstMRp!TpoRta
        psRegistro = psRegistro & s_Expresion & s_Caracter

         'nuevo 34: Modalidad del servicio prestado por el no domiciliado
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

         'nuevo 35: Aplicación del penultimo parrafo del Art. 76° de la Ley del Impuesto a la Renta
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter

        '41-36: Estado que identifica la oportunidad de la anotación o indicación si ésta corresponde a un ajuste.
'ini 2016-02-02.09  correccion ple

''        s_Expresion = Mid(Format(porstMRp!feedoc, "dd/mm/yyyy"), 4, 2)
''
''        Select Case porstMRp!codtdc
''              Case "05", "06", "07", "08", "11", "12", "13", "14"
''                   If s_Expresion = gsMesAct Then
''                      s_Expresion = "1"
''                   Else
''                      If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
''                         s_Expresion = "6"
''                      End If
''                      If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
''                         s_Expresion = "7"
''                      End If
''                   End If
''              Case Else
''                   If (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
''                      s_Expresion = "0"
''                   Else
''                      If s_Expresion = gsMesAct Then
''                         s_Expresion = "1"
''                      Else
''                         If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
''                            s_Expresion = "6"
''                         End If
''                         If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
''                            s_Expresion = "7"
''                         End If
''                      End If
''                   End If
''        End Select
'fin 2016-02-02.09  correccion ple
        
        '36: Estado que identifica la oportunidad de la anotación o indicación si ésta corresponde a un ajuste.
       s_Expresion = "0"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 37: Campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
fAE_08_2_reg_cpr = psRegistro

End Function

Private Function fAE_08_1_reg_cpr(xNroCorr As Long) As String
  Dim s_Expresion As String
  Dim psRegistro As String

        '+ 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 3: Nueva version Contribuyentes del Régimen General: Número
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 4: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 5: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        s_Expresion = IIf(Format(porstMRp!codtdc, "00") = "14", s_Expresion, "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 6: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 7: serie comprobante de pago
        's_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "", porstMRp!serdoc)
        s_Expresion = IIf(porstMRp!codtdc = "05", Right(porstMRp!serdoc, 1), porstMRp!serdoc) '2014-07-16
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 8: año emision DUA
        s_Expresion = IIf(IsNull(porstMRp!anno), "0", porstMRp!anno)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 9: numero comprobante de pago
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
        IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 10: numero final no dan derecho a credito - constante
        's_Expresion = "0"
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 11: tipo documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 12: numero documento identidad proveedor
        's_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        s_Expresion = IIf(IsNull(porstMRp!codaux), "", porstMRp!codaux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 13: razon social proveedor
        's_Expresion = IIf(IsNull(porstMRp!razAux), "-", porstMRp!razAux)
        s_Expresion = IIf(IsNull(porstMRp!razAux), "", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 14: base imponible adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp1), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 15: impuesto adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp2), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 16: base imponible adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp3), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 17: impuesto adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp4), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 18: base imponible adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp5), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 19: impuesto adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp6), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 20: adquisiciones no gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp7), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 21: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp8), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 22: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp9), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 23: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp10), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'NUEVO 24: Código  de la Moneda (Tabla 4)
        's_Expresion = IIf(porstMRp!tpomon = TPOMON_NAC, "PEN", "USD")
        s_Expresion = IIf(porstMRp!tpomon = TPOMON_NAC, CODMON_NAC, CODMON_EXT)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        
        '25-A 24: tipo de cambio
        'sololo mostrar si es diferen a N=tpomon
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!ImpTCb), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '26-A 25: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!feedoc_ref) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!feedoc_ref, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '27-A 26: tipo comprobante pago modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '28-A 27: serie comprobante pago modifica
        's_Expresion = IIf(IsNull(porstMRp!serdoc_ref), "-", porstMRp!serdoc_ref)
        s_Expresion = IIf(IsNull(porstMRp!serdoc_ref), "", porstMRp!serdoc_ref)
        's_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter

        '29-A 28: Nueva version Contribuyentes del Régimen General: Número
        s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '30-A 29: numero comprobante pago modifica
        's_Expresion = IIf(IsNull(porstMRp!nrodoc_ref), "-", porstMRp!nrodoc_ref)
        's_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        
        s_Expresion = IIf(IsNull(porstMRp!nrodoc_ref), "", porstMRp!nrodoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "", s_Expresion)
        '---------2016-02-11 correccion
        If s_Expresion = "" Then
            s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
            Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
            Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(s_Expresion, 7), _
            IIf(porstMRp!codtdc = "36", Right(s_Expresion, 8), s_Expresion))
        End If
        '---------
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2016-01-27 correccion ple
''        ' 30: numero comprobante pago no domiciliado
''        s_Expresion = IIf(IsNull(porstMRp!v1), "-", porstMRp!v1)
''        psRegistro = psRegistro & s_Expresion & s_Caracter
        'fin 2016-01-27 correccion ple
        
        '+ 31: fecha emision constancia detracción
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!FehCDt) Or s_Expresion = "0"), "01/01/0001", Format(porstMRp!FehCDt, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 32: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 33: comprobante afecto retencion
        's_Expresion = IIf(IsNull(porstMRp!indreten), "0", porstMRp!indreten)
        s_Expresion = IIf(IsNull(porstMRp!indreten), "", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 34: Clasificación de los bienes y servicios adquiridos
        s_Expresion = IIf(IsNull(porstMRp!TpoBns), "", porstMRp!TpoBns)
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
        'Nuevo 35: Identificación del Contrato o del proyecto.dejar vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
        'Nuevo 36: Error tipo 1: inconsistencia en el tipo de cambio.dejar vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 37: Error tipo 2: inconsistencia por proveedores no habidos.dejar vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 38: Error tipo 3: inconsistencia por proveedores.dejar vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 39: Error tipo 4: inconsistencia por DNIs .dejar vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
        'Nuevo 40: Indicador de Comprobantes de pago cancelados. siempre poner valor 1
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
        '41-34: identifica ajuste - constante
        s_Expresion = Mid(Format(porstMRp!feedoc, "dd/mm/yyyy"), 4, 2)
        
        Select Case porstMRp!codtdc
              Case "05", "06", "07", "08", "11", "12", "13", "14"
                   If s_Expresion = gsMesAct Then
                      s_Expresion = "1"
                   Else
                      If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                         s_Expresion = "6"
                      End If
                      If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                         s_Expresion = "7"
                      End If
                   End If
              Case Else
                   If (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
                      s_Expresion = "0"
                   Else
                      If s_Expresion = gsMesAct Then
                         s_Expresion = "1"
                      Else
                         If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                            s_Expresion = "6"
                         End If
                         If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                            s_Expresion = "7"
                         End If
                      End If
                   End If
        End Select
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 34: Campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
    fAE_08_1_reg_cpr = psRegistro
End Function
Private Function fAE_14_1_reg_vta(xNroCorr As Long) As String

  Dim s_Expresion As String
  Dim psRegistro As String

        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actuliza libro electronico
        ' 3: Version nueva
'''        'Contribuyentes del Régimen General: Número correlativo del asiento contable
'''        'identificado en el campo 2, cuando se utilice el Código Único de la Operación
'''        '(CUO). El primer dígito debe ser: "A" para el asiento de apertura del
'''        'ejercicio, "M" para los asientos de movimientos o ajustes del mes o "C" para
'''        'el asiento de cierre del ejercicio.
'''        '2. Contribuyentes del Régimen Especial de Renta - RER:  Número correlativo.
'''        'El primer dígito debe ser: "M".
        's_Expresion = "M" & porstMRp!NroCpb
        '2014-07-31 numero correlativo s_Expresion = "M" & Left(porstMRp!NroCpb, 4) & Right(porstMRp!NroCpb, 5)
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'fin 2015-05-23 actuliza libro electronico
              
        ' 4: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 5: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 6: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 7: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "", porstMRp!serdoc)
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 8: numero comprobante de pago
        'ini 2015-05-26 actualiza libro electronico
        's_Expresion = porstMRp!nrodoc
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), porstMRp!nrodoc)
        'fin 2015-05-26 actualiza libro electronico
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 9: numero final agrupar documentos
'ini 2016-02-02.09  correccion ple
'        s_Expresion = IIf(IsNull(porstMRp!nrodoc_fin), "0", porstMRp!nrodoc_fin)
'        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
'fin 2016-02-02.09  correccion ple
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 10: tipo documento identidad cliente
        s_Expresion = Right(IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci), 1)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 11: numero documento identidad cliente
'ini 2014-07-10 validacion ruc 8dig tdoc=01
        's_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        If IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci) = "01" Then
             s_Expresion = IIf(IsNull(porstMRp!codaux), "", porstMRp!codaux)
             s_Expresion = Right(s_Expresion, 8)
       Else
            s_Expresion = IIf(IsNull(porstMRp!codaux), "", porstMRp!codaux)
        End If
'ini 2014-07-10 validacion ruc 8dig tdoc=01
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 12: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 13: valor facturado exportacion
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impexp), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 14: base imponible operacion gravada
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impogr), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2016-02-02.05 correccion ple

        'Nuevo 15: Descuento de la Base Imponible.deja vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '16- 18: igv y/o ipm
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impigv), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 17: deja vacio.Descuento del Impuesto General a las Ventas y/o Impuesto de Promoción Municipal
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2016-02-02.05 correccion ple
        
        '18- 15: importe total operacion exonerada
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impexo), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '19- 16: importe total operacion inafecta
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impina), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '20- 17: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impisc), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'ini 2016-02-02.05 correccion ple
'        ' 18: igv y/o ipm
'        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impigv), 2), ",", "")
'        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2016-02-02.05 correccion ple

        '21- 19: base imponible operacion gravada ivap - constante
        s_Expresion = "0.00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '22- 20: impuesto ventas arroz pilado (ivap) - constante
        s_Expresion = "0.00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '23- 21: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impoim), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '24- 22: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imptot), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'ini 2016-02-02.05 correccion ple
         'NUEVO 25: Código  de la Moneda (Tabla 4)
        's_Expresion = IIf(porstMRp!tpomon = TPOMON_NAC, "PEN", "USD")
        
        s_Expresion = porstMRp!codmon '2016-02-02.06  correccion ple
        '2016-02-09 teo dijo provisionalmente en blanco s_Expresion = "" '2016-02-02.06  correccion ple
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2016-02-02.05 correccion ple
               
        '26- 23: tipo de cambio
        s_Expresion = Replace(FormatNumber(CDec(IIf(IsNull(porstMRp!ImpTCb), 0, porstMRp!ImpTCb)), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '27- 24: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!d2), "0", porstMRp!d2), "00")
        s_Expresion = IIf((IsNull(porstMRp!d1) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!d1, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '28- 25: tipo comprobante pago modifica - codtdc_ref
        s_Expresion = Format(IIf(IsNull(porstMRp!d2), "0", porstMRp!d2), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '29- 26: serie comprobante pago modifica- serdoc_ref
        s_Expresion = IIf(IsNull(porstMRp!d3), "", porstMRp!d3)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '30- 27: numero comprobante pago modifica- nrodoc_ref
        s_Expresion = IIf(IsNull(porstMRp!d4), "00", porstMRp!d4)
        
        '2016-02-11
        If s_Expresion <> "00" Then
            s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
            Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
            Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(s_Expresion, 7), s_Expresion)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
'ini 2016-02-02.05 correccion ple
''        'ini 2015-05-23 actuliza libro electronico
''        ' 28: version nueva  Valor FOB embarcado de la exportación
''        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impfob_mn), 2), ",", "")
''        psRegistro = psRegistro & s_Expresion & s_Caracter
''        'fin 2015-05-23 actuliza libro electronico
'fin 2016-02-02.05 correccion ple
        
        'Nuevo 31:Identificación del Contrato o del proyecto en el caso de los Operadores de las sociedades.deja vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 32: Error tipo 1: inconsistencia en el tipo de cambio.deja vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 33: Indicador de Comprobantes de pago cancelados con medios de pago (Tabla 1).deja vacio
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
        '34- 29: identifica estado comprobante periodo - constante
        s_Expresion = IIf(CDec(porstMRp!imptot) = 0, "2", "1")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '35- 29: campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
        fAE_14_1_reg_vta = psRegistro
End Function
Private Function fAE_14_1_reg_vta_sql(sMoneda As String) As String
    Dim sSentencia As String
    sSentencia = "SELECT concat(vta.coddro,vta.nrocpb) as nrocpb, date_format(vta.feedoc,'%d/%m/%Y') as feedoc, date_format(vta.fevdoc,'%d/%m/%Y') as fevdoc, vta.codtdc as codtdc, vta.serdoc as serdoc, "
    sSentencia = sSentencia & "vta.nrodoc, vta.nrodoc_fin, aux.tpodci as tpodci, (CASE WHEN tpodci='01' THEN RIGHT(aux.rucaux, 8) ELSE aux.rucaux END) as codaux, trim(left(aux.razaux,60)) as razaux, "
    sSentencia = sSentencia & "ROUND(vta.impexp_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impexp, "
    sSentencia = sSentencia & "ROUND(vta.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impogr, "
    sSentencia = sSentencia & "(CASE WHEN vta.categoriadoc<>'" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impexo, "
    sSentencia = sSentencia & "(CASE WHEN vta.categoriadoc='" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impina, "
    sSentencia = sSentencia & "ROUND(vta.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impisc, "
    sSentencia = sSentencia & "ROUND(vta.impigv_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impigv, "
    sSentencia = sSentencia & "ROUND(vta.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impoim, "
    sSentencia = sSentencia & "ROUND(vta.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as imptot, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "vta.imptcb as imptcb, "
    'por teo 2016-02-10 sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN vta.imptcb ELSE 0 END imptcb, "
    sSentencia = sSentencia & "vta.imptcb as imptcb, "
'fin 2015-09-16 t.cam solo ME
    sSentencia = sSentencia & "date_format(feedoc_ref,'%d/%m/%Y') as d1, codtdc_ref as d2, serdoc_ref as d3, nrodoc_ref as d4, '" & sMoneda & "' as Moneda "
    'ini 2014-05-25
    sSentencia = sSentencia & ",impfob_mn "
    'fin 2014-05-25
    'sSentencia = sSentencia & ",tpomon,CodMon " '2016-02-02.03 correccion ple
    sSentencia = sSentencia & ",tpomon,TRIM(" & fIsNull() & "CodMon,'')) CodMon " '2016-02-02.03 correccion ple
    sSentencia = sSentencia & "FROM covtadoc vta "
    sSentencia = sSentencia & "INNER JOIN tgaux aux on vta.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc on vta.codemp=tdc.codemp and vta.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE vta.codemp='" & gsCodEmp & "' and vta.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    sSentencia = sSentencia & "ORDER BY vta.codtdc,vta.serdoc,vta.nrodoc"
    fAE_14_1_reg_vta_sql = sSentencia
End Function
Private Function fAE_05_1_lib_diario(xNroCorr As Long) As String
  Dim s_Expresion As String
  Dim psRegistro As String
  Dim xxMesPvs As String
  xxMesPvs = Left(cmbEjercicio.Text, 2)
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double

        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        psRegistro = psRegistro & s_Expresion & s_Caracter
         
         ' 3: segun nueva estructura numero correlativo de asiento contable
         s_Expresion = gfCeros(Str(xNroCorr), 8, 0, "0")
         If ExistFieldInRS(porstMRp, "mespvs") Then
            psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         Else
            psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         End If
         
'ini 2016-02-02.07  correccion ple
''         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
''         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
'fin 2016-02-02.07  correccion ple
    
        '4-5: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 5: Código de la Unidad de Operación, de la Unidad Económica Administrativa
        s_Expresion = "" 'dejar vacio
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 6: Código del Centro de Costos, Centro de Utilidades
        s_Expresion = porstMRp!codcco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 7: Tipo de Moneda de origen (Tabla 4)
        s_Expresion = porstMRp!codmon
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 8: Tipo de documento de identidad del emisor  (Tabla 2)
'ini 2016-03-22 error diario y mayor ple solu
'        s_Expresion = porstMRp!TpoDci
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2016-03-22 error diario y mayor ple solu
        
        'Nuevo 9: Tipo de documento de identidad del emisor  (Tabla 2)
'ini 2016-03-22 error diario y mayor ple solu
'        s_Expresion = porstMRp!codaux
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2016-03-22 error diario y mayor ple solu
        
        'Nuevo 10: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 11: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 12: numero comprobante de pago
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
        IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 13: Fecha contable
        s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        
        'Nuevo 14: Fecha de vencimiento
        s_Expresion = Format(porstMRp!fevdoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '15-6: Fecha de la operación o emisión
        's_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy") '2016-02-02.07  correccion ple
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '16-7: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 17: Glosa referencial, de ser el caso
        '2016-03-01 s_Expresion = porstMRp!codaux
        s_Expresion = porstMRp!RefDoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        
        '18- 8: movimiento debe
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '19- 9: movimiento haber
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
                
        'Nuevo 20: Dato Estructurado: Código del libro, campo 1, campo 2 y campo 3 del Registro de Ventas e Ingresos
        s_Expresion = "" 'dejar vacio
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        
'ini 2016-02-02.07  correccion ple
''        '10: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
''        psRegistro = psRegistro & s_Caracter
''
''        '11: segun nueva estructura Número correlativo utilizado en el Registro de Compras
''        psRegistro = psRegistro & s_Caracter
''
''        '12: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
''        psRegistro = psRegistro & s_Caracter
'fin 2016-02-02.07  correccion ple
                
        '21- 13: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '22- 14: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '23- 15: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '24- 16: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '25- 17: codigo libro - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codlib) Then
          s_Expresion = porstMRp!codlib
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    fAE_05_1_lib_diario = psRegistro
End Function

Private Function fAE_05_1_lib_diario_sql(sMoneda As String) As String
    Dim sSentencia As String

    sSentencia = "SELECT det.CodDro, det.nrocpb, det.nroite, det.fehope,"
    sSentencia = sSentencia & fIsNull() & "det.codtdc,'') codtdc," & fIsNull() & "det.serdoc,'') serdoc," & fIsNull() & "det.nrodoc,'') nrodoc, "
    sSentencia = sSentencia & "det.codcta, cta.detcta," & fIsNull() & " det.codaux,'') codaux," & fIsNull() & " aux.razaux,'') razaux,"
    '2016-03-01 sSentencia = sSentencia & "det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & fIsNull() & "det.refdoc,'') refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & "tdc.AbvTDc, " & Choose(gsIdioma, "dro.DetDro", "dro.DetDrox") & " AS DetDro, "
    sSentencia = sSentencia & "'" & sMoneda & "' as sMoneda, dro.codlib "
    sSentencia = sSentencia & ",det.mespvs " '2015-04-06 convierte un solo mes 12y00
    sSentencia = sSentencia & "," & fIsNull() & "det.codmon,'') codmon ," & fIsNull() & "right(aux.tpodci,1),'') as tpodci,"  '2016-02-02.07  correccion ple
    sSentencia = sSentencia & fIsNull() & "det.codcco,'') codcco ,fevdoc,feedoc," & fIsNull() & "det.refdoc,'') refdoc "  '2016-02-02.07  correccion ple
    sSentencia = sSentencia & "FROM (((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
'ini 2015-04-01 convierte un solo mes 12y00
    'sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    If xqmes = "12" Then
        sSentencia = sSentencia & "AND det.mespvs >='" & xqmes & "'" & " AND det.mespvs <='13' "
    ElseIf xqmes = "01" Then
'ini 2015-04-20 corre 00 en ene dia
        sSentencia = sSentencia & "AND det.mespvs >='00'" & " AND det.mespvs <='" & xqmes & "' "
'fin 2015-04-20 corre 00 en ene dia
    Else
    sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    End If
'fin 2015-04-01 convierte un solo mes 12y00
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    '2015-04-01 convierte un solo mes 12y00 sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite, fehope"
    sSentencia = sSentencia & "ORDER BY mespvs, coddro, nrocpb, nroite, fehope"
fAE_05_1_lib_diario_sql = sSentencia
End Function
Private Function fAE_05_2_lib_diario_simplificado(xNroCorr As Long) As String
  Dim s_Expresion As String
  Dim psRegistro As String
  Dim xxMesPvs As String
  xxMesPvs = Left(cmbEjercicio.Text, 2)
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double
       ' 1: periodo
'ini 2015-04-01 convierte un solo mes 12y00
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        If Left(cmbEjercicio.Text, 2) = "12" Then
'        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        Else
'        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        End If
'fin 2015-04-01 convierte un solo mes 12y00
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
         
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
         ' 3: segun nueva estructura numero correlativo de asiento contable
         '2014-07-31 numero correlativo s_Expresion = porstMRp!coddro & Right(Format(porstMRp!NroCpb, "000000"), 5)
         s_Expresion = gfCeros(Str(xNroCorr), 8, 0, "0")
         '2015-04-06 convierte un solo mes 12y00  psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         If ExistFieldInRS(porstMRp, "mespvs") Then
            psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         Else
            psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         End If
         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
    
        ' 3: plan de cuentas - constante
'  2014-05-29 sale re`plaza por gnCodPlaCata        s_Expresion = "01"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: fecha emisión
        s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        ' 7: movimiento debe
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: movimiento haber
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
        ' 11: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
        psRegistro = psRegistro & s_Caracter
        ' 12: segun nueva estructura Número correlativo utilizado en el Registro de Compras
        psRegistro = psRegistro & s_Caracter
        ' 13: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
        psRegistro = psRegistro & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
                
        ' 9: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: codigo libro - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codlib) Then
          s_Expresion = porstMRp!codlib
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'******************************************************************
fAE_05_2_lib_diario_simplificado = psRegistro
End Function
Private Function fAE_05_2_lib_diario_simplificado_sql(sMoneda As String) As String
    Dim sSentencia As String
    sSentencia = "SELECT det.CodDro, det.nrocpb, det.nroite, det.fehope, det.codtdc, det.serdoc, det.nrodoc, "
    sSentencia = sSentencia & "det.codcta, cta.detcta, det.codaux, aux.razaux, det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & "tdc.AbvTDc, " & Choose(gsIdioma, "dro.DetDro", "dro.DetDrox") & " AS DetDro, "
    sSentencia = sSentencia & "'" & sMoneda & "' as sMoneda, dro.codlib "
    sSentencia = sSentencia & "FROM (((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite, fehope"

fAE_05_2_lib_diario_simplificado_sql = sSentencia
End Function

Private Function fAE_05_3_lib_diario_deta_plan_cta(xNroCorr As Long) As String
  Dim s_Expresion As String
  Dim psRegistro As String
       ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & Format(Day(gfUltDia("01/" & Left(cmbEjercicio.Text, 2) & "/" & gsAnoAct)), "00") ' Format(Str(Day(Now)), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: Código de la Cuenta Contable desagregada hasta el nivel máximo de dígitos utilizado
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 3: Descripción de la Cuenta Contable desagregada al nivel máximo de dígitos utilizado
        s_Expresion = porstMRp!detcta
        psRegistro = psRegistro & s_Expresion & s_Caracter

        ' 4: Código del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = gnCodPlaCata
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 5: Descripción del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = IIf(gnCodPlaCata <> "99", "-", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 6: Código del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nuevo 7: Código del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
        '8- 6: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter

fAE_05_3_lib_diario_deta_plan_cta = psRegistro
End Function
Private Function fAE_05_3_lib_diario_deta_plan_cta_sql(sMoneda As String) As String
    Dim sSentencia As String
    'gsNivCta
    Dim xxniv1 As String
    Dim xxnivf As String
    xxniv1 = Left(gsNivCta, 1)
    xxnivf = Right(gsNivCta, 1)
    sSentencia = "SELECT"
    sSentencia = sSentencia & "      a.codcta , a.detcta, b.codcta codcta2, b.detcta detcta2 "
    sSentencia = sSentencia & "FROM cocta a "
    sSentencia = sSentencia & "LEFT JOIN cocta b "
    sSentencia = sSentencia & "    ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.codcta," & xxniv1 & ")=LEFT(b.codcta," & xxniv1 & ") "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "    AND a.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "    AND LENGTH(a.codcta)=" & xxnivf
    sSentencia = sSentencia & "    AND  LENGTH(b.codcta)=" & xxniv1 & " "
    sSentencia = sSentencia & "ORDER BY codcta"
'fin 2014-05-30 adicion 5.3 plan ctas

fAE_05_3_lib_diario_deta_plan_cta_sql = sSentencia
End Function

Private Function fAE_06_1_lib_mayor(xNroCorr As Long) As String
  Dim s_Expresion As String
  Dim psRegistro As String
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double

        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
         psRegistro = psRegistro & s_Expresion & s_Caracter
         
         ' 3: segun nueva estructura numero correlativo de asiento contable
         s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
         psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         
'ini 2016-02-02.07  correccion ple
''         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
''         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
'fin 2016-02-02.07  correccion ple
    
        '4- 5: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        
        'Nuevo 5: Código de la Unidad de Operación, de la Unidad Económica Administrativa
        s_Expresion = "" 'dejar vacio
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 6: Código del Centro de Costos, Centro de Utilidades
        s_Expresion = porstMRp!codcco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 7: Tipo de Moneda de origen (Tabla 4)
        s_Expresion = porstMRp!codmon
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 8: Tipo de documento de identidad del emisor  (Tabla 2)
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 9: Tipo de documento de identidad del emisor  (Tabla 2)
        s_Expresion = porstMRp!codaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 10: tipo comprobante de pago
'ini 2016-03-22 error diario y mayor ple solu
        's_Expresion = Format(porstMRp!codtdc, "00")
        s_Expresion = IIf(Trim(porstMRp!codtdc) = "", "00", Format(porstMRp!codtdc, "00"))
'fin 2016-03-22 error diario y mayor ple solu
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 11: serie comprobante de pago
'ini 2016-03-22 error diario y mayor ple solu
        's_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        s_Expresion = IIf(IsNull(porstMRp!serdoc) Or Trim(porstMRp!serdoc) = "", "-", porstMRp!serdoc)
'fin 2016-03-22 error diario y mayor ple solu
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 12: numero comprobante de pago
'ini 2016-03-22 error diario y mayor ple solu
'        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
'        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
'        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
'        IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        If Trim(porstMRp!nrodoc) = "" Then
            s_Expresion = "-"
        Else
            s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
            Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
            Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
            IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        End If
'fin 2016-03-22 error diario y mayor ple solu
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 13: Fecha contable
        s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        
        'Nuevo 14: Fecha de vencimiento
        s_Expresion = Format(porstMRp!fevdoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
                                
        '15- 6: fecha emisión
'ini 2016-02-02.07  correccion ple
''        s_Expresion = "01/" & Left(cmbEjercicio.Text, 2) & "/" & gsAnoAct
''        If Not IsNull(porstMRp!fehope) Then
''          s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
''        End If
'fin 2016-02-02.07  correccion ple
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '16-7: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Nuevo 17: Glosa referencial, de ser el caso
        '2016-03-01 s_Expresion = porstMRp!codaux
        s_Expresion = porstMRp!RefDoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        '18- 8: saldo o movimiento deudor
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '19- 9: saldo o movimiento acreedor
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
                 
        'Nuevo 20: Dato Estructurado: Código del libro, campo 1, campo 2 y campo 3 del Registro de Ventas e Ingresos
        s_Expresion = "" 'dejar vacio
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
'ini 2016-02-02.07  correccion ple
'''        '10: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
'''        psRegistro = psRegistro & s_Caracter
'''        '11: segun nueva estructura Número correlativo utilizado en el Registro de Compras
'''        psRegistro = psRegistro & s_Caracter
'''        '12: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
'''        psRegistro = psRegistro & s_Caracter
'fin 2016-02-02.07  correccion ple
        
        '21- 13: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '22- 14: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '23- 15: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '24- 16: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '25- 17: medio pago - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!medpago) Then
          s_Expresion = porstMRp!medpago
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter

fAE_06_1_lib_mayor = psRegistro
End Function

Private Function fAE_06_1_lib_mayor_sql(sMoneda As String) As String
    Dim sSentencia As String
    Dim s_MesIni As String, s_MesFin As String
    Dim sSalAntDeb As String, sSalAntHab As String
    Dim nSecuencia As Long

    s_MesIni = Left(cmbEjercicio.Text, 2)
    s_MesFin = Left(cmbEjercicio.Text, 2)
    ' movimientos
    '2015-04-07 error mespvs sSentencia = "SELECT  det.mespvs,det.codcta, det.coddro, det.nrocpb, det.nroite, det.fehope, "
    'sSentencia = "SELECT  CASE det.mespvs WHEN det.mespvs>='12' THEN det.mespvs ELSE " & s_MesIni & " END mespvs, "
    sSentencia = "SELECT " & IIf(s_MesIni = "12", "det.mespvs", "'" & s_MesIni & "'") & " as mespvs, "
    sSentencia = sSentencia & " det.codcta, det.coddro, det.nrocpb, det.nroite, det.fehope, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-' + det.serdoc + '-' + det.nrodoc)") & " AS cDocume, "
'2016-02-02.07  correccion plesSentencia = sSentencia & "det.codaux, aux.RazAux, det.refdoc, det.tpodoc AS medpago, "
'2016-03-01 sSentencia = sSentencia & fIsNull() & " det.codaux,'') codaux," & fIsNull() & " aux.razaux,'') razaux, det.refdoc, det.tpodoc AS medpago, "
    sSentencia = sSentencia & fIsNull() & " det.codaux,'') codaux," & fIsNull() & " aux.razaux,'') razaux, " & fIsNull() & "det.refdoc,'') refdoc, det.tpodoc AS medpago, "
    sSentencia = sSentencia & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, " & Choose(gsIdioma, "dro.detdro", "dro.detdrox") & " AS detdro, tdc.abvtdc "
    'sSentencia = sSentencia & ",det.mespvs " '2015-04-06 convierte un solo mes 12y00
    sSentencia = sSentencia & "," & fIsNull() & "det.codmon,'') codmon ," & fIsNull() & "right(aux.tpodci,1),'') as tpodci,"  '2016-02-02.07  correccion ple
    '2016-03-01 sSentencia = sSentencia & fIsNull() & "det.codcco,'') codcco ,fevdoc,feedoc,det.refdoc "  '2016-02-02.07  correccion ple
    sSentencia = sSentencia & fIsNull() & "det.codcco,'') codcco ,fevdoc,feedoc "  '2016-02-02.07  correccion ple
    sSentencia = sSentencia & "," & fIsNull() & "det.codtdc,'') codtdc," & fIsNull() & "det.serdoc,'') serdoc," & fIsNull() & "det.nrodoc,'') nrodoc " '2016-02-02.07  correccion ple
    sSentencia = sSentencia & "FROM ((((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
'ini 2015-04-01 convierte un solo mes 12y00
    'sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & s_MesFin & "' "
    If s_MesIni = "12" Then
    sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & "13" & "' "
    Else
    sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & s_MesFin & "' "
    End If
'fin 2015-04-01 convierte un solo mes 12y00
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    ' Saldo anterior
    '2015-04-20 solo sdo ante de ene a dic If s_MesIni <> "00" Then
    'If (s_MesIni <> "00") Or Not (s_MesIni >= "02" And s_MesIni <= "12") Then
    If s_MesIni = "01" Or s_MesIni = "13" Then
      sSalAntDeb = "ROUND(("
      sSalAntHab = "ROUND(("
      For nSecuencia = 0 To (Val(s_MesIni) - 1)
        sSalAntDeb = sSalAntDeb & "acu.acud" & Format(nSecuencia, "00") & "_" & sMoneda & IIf(nSecuencia = (Val(s_MesIni) - 1), ")", "+")
        sSalAntHab = sSalAntHab & "acu.acuh" & Format(nSecuencia, "00") & "_" & sMoneda & IIf(nSecuencia = (Val(s_MesIni) - 1), ")", "+")
      Next nSecuencia
      sSalAntDeb = sSalAntDeb & ", 2)"
      sSalAntHab = sSalAntHab & ", 2)"
      '2015-04-20 corre 00 en ene dia
      sSentencia = sSentencia & "UNION "
      '2015-04-07 error mespvs sSentencia = sSentencia & "SELECT '00' AS mespvs, cta.codcta, Null, Null, Null, Null, Null, Null, Null, Null, Null, "
      'sSentencia = sSentencia & "SELECT " & IIf(s_MesIni = "12", "'" & s_MesIni & "'", "'00'") & " AS mespvs, "
      sSentencia = sSentencia & "SELECT " & "'" & s_MesIni & "'" & " AS mespvs, "
      sSentencia = sSentencia & " cta.codcta, Null, Null, Null, Null, Null, Null, Null, Null, Null, "
      sSentencia = sSentencia & "'" & Choose(gsIdioma, "SALDO ANTERIOR", "PREVIOUS BALANCE") & "' AS gloite, "
      sSentencia = sSentencia & sSalAntDeb & " AS nDebe, "
      sSentencia = sSentencia & sSalAntHab & " AS nHaber, "
      sSentencia = sSentencia & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, Null, Null "
      sSentencia = sSentencia & ",Null ,Null " '2016-02-02.07  correccion ple codmon
      '2016-03-01 sSentencia = sSentencia & ",Null ,Null,Null ,Null " '2016-02-02.07  correccion ple codmon
      sSentencia = sSentencia & ",Null ,Null,Null  " '2016-02-02.07  correccion ple codmon
      sSentencia = sSentencia & ",Null ,Null,Null " '2016-02-02.07  correccion ple codmon
      sSentencia = sSentencia & "FROM cocta cta "
      sSentencia = sSentencia & "LEFT JOIN coctaacu acu ON acu.codemp=cta.codemp AND acu.pdoano=cta.pdoano AND acu.codcta=cta.codcta "
      sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND cta.tpocta='" & TPOCTA_TRA & "' "
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & "HAVING (ROUND((nDebe-nHaber), 2)<>0.00) "
      Else
        sSentencia = sSentencia & "AND (ROUND((" & sSalAntDeb & "-" & sSalAntHab & "), 2)<>0.00) "
      End If
      '2015-04-20 corre 00 en ene dia
      
    End If
    sSentencia = sSentencia & "ORDER BY codcta, mespvs, coddro, nrocpb, nroite"
    fAE_06_1_lib_mayor_sql = sSentencia

End Function

Private Sub pAE_08_11_reg_cpr_2016_01_27()
        '+ 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actualiza libro electronico
        '+ 3: Nueva version Contribuyentes del Régimen General: Número
        '2014-07-31 numero correlativo s_Expresion = "M" & Left(porstMRp!NroCpb, 4) & Right(porstMRp!NroCpb, 5)
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'ini 2015-05-23 actuliza libro electronico
        
        '+ 4: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '+ 5: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        s_Expresion = IIf(Format(porstMRp!codtdc, "00") = "14", s_Expresion, "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 6: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 7: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        s_Expresion = IIf(porstMRp!codtdc = "05", Right(porstMRp!serdoc, 1), porstMRp!serdoc) '2014-07-16
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 8: año emision DUA
        s_Expresion = IIf(IsNull(porstMRp!anno), "0", porstMRp!anno)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 9: numero comprobante de pago
        'ini 2015-05-26 actualiza libro electronico
        's_Expresion = porstMRp!nrodoc
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
        IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        '2014-08-20 Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), porstMRp!nrodoc)
        'fin 2015-05-26 actualiza libro electronico
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 10: numero final no dan derecho a credito - constante
        s_Expresion = "0"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 11: tipo documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 12: numero documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 13: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "-", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 14: base imponible adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp1), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 15: impuesto adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp2), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 16: base imponible adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp3), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 17: impuesto adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp4), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 18: base imponible adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp5), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 19: impuesto adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp6), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 20: adquisiciones no gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp7), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 21: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp8), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 22: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp9), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 23: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp10), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 24: tipo de cambio
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!ImpTCb), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 25: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!feedoc_ref) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!feedoc_ref, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 26: tipo comprobante pago modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 27: serie comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!serdoc_ref), "-", porstMRp!serdoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
'      txtDetalle(0).MaxLength = .uorstMain!codaduana.DefinedSize
'      txtDetalle(1).MaxLength = .uorstMain!annodua.DefinedSize
'      txtDetalle(2).MaxLength = .uorstMain!nrodua.DefinedSize

        'ini 2015-05-23 actuliza libro electronico
        ' 28: Nueva version Contribuyentes del Régimen General: Número
        '2014-07-10 cambiar s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana) & IIf(IsNull(porstMRp!annodua), "", porstMRp!annodua) & IIf(IsNull(porstMRp!nrodua), "", porstMRp!nrodua)
        s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'ini 2015-05-23 actuliza libro electronico
        
        ' 29: numero comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!nrodoc_ref), "-", porstMRp!nrodoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 30: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!v1), "-", porstMRp!v1)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 31: fecha emision constancia detracción
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!FehCDt) Or s_Expresion = "0"), "01/01/0001", Format(porstMRp!FehCDt, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 32: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 33: comprobante afecto retencion
        s_Expresion = IIf(IsNull(porstMRp!indreten), "0", porstMRp!indreten)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '41-34: identifica ajuste - constante
        s_Expresion = Mid(Format(porstMRp!feedoc, "dd/mm/yyyy"), 4, 2)
        
        '2015-02-19 s_Expresion = IIf(s_Expresion = gsMesAct, "1", "6")
        's_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
        
        'ini 215-08-10 adicionado por teo
'        If (porstMRp!codtdc = "05" Or porstMRp!codtdc = "12") And (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "1", IIf(s_Expresion = gsMesAct, "1", "6"))
'        Else
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
'        End If

'        If (porstMRp!codtdc = "05" Or porstMRp!codtdc = "12") And (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "1", IIf(s_Expresion = gsMesAct, "1", "6"))
'        Else
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
'        End If
'        'ini 2015-08-10 correcion rafael fope-femision
'        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 12 Then
'            s_Expresion = "6"
'        End If
'        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 12 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
'            s_Expresion = "7"
'        End If
'ini 2015-08-26
          Select Case porstMRp!codtdc
                Case "05", "06", "07", "08", "11", "12", "13", "14"
                     If s_Expresion = gsMesAct Then
                        s_Expresion = "1"
                     Else
                        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                           s_Expresion = "6"
                        End If
                        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                           s_Expresion = "7"
                        End If
                     End If
                Case Else
                     If (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
                        s_Expresion = "0"
                     Else
                        If s_Expresion = gsMesAct Then
                           s_Expresion = "1"
                        Else
                           If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                              s_Expresion = "6"
                           End If
                           If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                              s_Expresion = "7"
                           End If
                        End If
                     End If
          End Select
'fin 2015-08-26
      

        
       'fin 2015-08-10 correcion rafael fope-femision
       'fin 215-08-10 adicionado por teo
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 34: Campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
End Sub

Private Sub ppRegistroElectronico()

  'ini rcs 2015-04-27 correccion version
        Dim arrDocume(4) As String
        arrDocume(0) = "d.AbvTDc"
        arrDocume(1) = "'-'"
        arrDocume(2) = "a.SerDoc"
        arrDocume(3) = "'-'"
        arrDocume(4) = "a.NroDoc"
  'ini rcs 2015-04-27 correccion version

  Dim sSentencia As String, sArchivo As String
  Dim sMoneda As String, sExpresion As String
  Dim s_MesIni As String, s_MesFin As String
  Dim sSalAntDeb As String, sSalAntHab As String
  Dim nSecuencia As Long

  ' valido información
  If pnOpcion = 99 Then MsgBox "Seleccionar Libro o Registro", vbCritical, "Sistema Contable": Exit Sub
  If cboTpoMon.Text = "" Then MsgBox "Seleccionar Moneda", vbCritical, "Sistema Contable": Exit Sub
  
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "MN", "ME")
  xqmes = Left(cmbEjercicio.Text, 2)
  Select Case pnOpcion
   'Case 5: sSentencia = LibroCajaElectro(999): sArchivo = "050100" & "00" & "2": Exit Sub '2015-04-23 nvos.rpt teo 3.2 LIBRO DE INVENTARIOS Y BALANCES
   Case 5: sSentencia = LibroCajaElectro(999): sArchivo = "030100" & "00" & "2" '2015-04-23 nvos.rpt teo 3.2 LIBRO DE INVENTARIOS Y BALANCES
   
'ini 2015-11-12/18 PLE rpt fmt electro archivo

'ini 2015-12-04 libros inven y balance correcc
'   Case 6, 7, 8, 9, 10
'        '2015-11-12/18 PLE rpt fmt electro archivo sArchivo = "040100" & "00" & "1"
'        Select Case pnOpcion
'        Case 6, 7: sArchivo = "040100" & "00" & "1"
'        Case 8: sArchivo = "030400" & "00" & "1"
'        Case 9, 10: sArchivo = "030500" & "00" & "1"
'        Case Else: sArchivo = "sin_archivo"
'        End Select
'        sSentencia = ""
'        sSentencia = sSentencia & "SELECT"
'        sSentencia = sSentencia & "    MIN(mespvs) mespvs,"
'        sSentencia = sSentencia & "    MIN(a.coddro) coddro,MIN(a.nrocpb) nrocpb,MIN(a.nroite) nroite,"
'        sSentencia = sSentencia & fIsNull() & "MAX(c.tpodci),'') tpodci,"
'        sSentencia = sSentencia & "    a.CodAux,"
'        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
'        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
'        sSentencia = sSentencia & "    MIN(a.FeEDoc) AS FeEDoc,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
'        sSentencia = sSentencia & "From (COCpbDet a "
'        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
'        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
'        '2015-11-12/18 PLE rpt fmt electro archivo sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
'        Select Case pnOpcion
'        Case 6, 7
'        sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
'        Case 8
'        sSentencia = sSentencia & " AND left(a.CodCta,2)='14'  "
'        Case 9, 10
'        sSentencia = sSentencia & " AND (left(a.CodCta,2)='16'  OR left(a.CodCta,2)='17')  "
'        Case Else
'        sSentencia = sSentencia & " "
'        End Select
'
'        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
'        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux "
'        'sSentencia = sSentencia & "Having (Round(DebeSol - HaberSol, 2) <> 0.00 Or Round(DebeDol - HaberDol, 2) <> 0.00) "
'        If ps_Plataforma = pSrvMySql Then
'            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
'        Else
'            sSentencia = sSentencia & " HAVING "
'            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
'            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
'        End If
'        sSentencia = sSentencia & "ORDER BY a.CodAux"
'
'case=6y7 3.3 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 12 CUENTAS POR COBRAR COMERCIALES – TERCEROS Y 13
'case=8 3.4 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS
'case=9y10 3.5 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS - TERCEROS O CUENTA 17
'case=11 3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
'case=12 3.11 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 41 REMUNERACIONES Y PARTICIPACIONES POR PAGAR (PCGE) (2)
'case=13y14 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=15y16 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=18 3.15  LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 37 ACTIVO DIFERIDO Y DE LA CUENTA 49 PASIVO DIFERIDO (PCGE)   (2)
'archivo agrupado y ordenado por RUC
   Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18
        '2015-11-12/18 PLE rpt fmt electro archivo sArchivo = "040100" & "00" & "1"
        Select Case pnOpcion
        Case 6, 7: sArchivo = "040100" & "00" & "1"
        Case 8: sArchivo = "030400" & "00" & "1"
        Case 9, 10: sArchivo = "030500" & "00" & "1"
        Case 11: sArchivo = "030600" & "00" & "1"
        Case 12: sArchivo = "031100" & "00" & "1"
        Case 13, 14: sArchivo = "041300" & "00" & "1"
        Case 15, 16: sArchivo = "041400" & "00" & "1"
        Case 18: sArchivo = "031500" & "00" & "1"
        Case Else: sArchivo = "sin_archivo"
        End Select
        sSentencia = ""
        sSentencia = sSentencia & "SELECT"
        sSentencia = sSentencia & "    a.mespvs,"
        sSentencia = sSentencia & "    a.coddro, a.nrocpb, a.nroite,"
        sSentencia = sSentencia & fIsNull() & "c.tpodci,'') tpodci,"
'ini 2015-12-07 libros inven y balance correcc
        'sSentencia = sSentencia & "    a.CodAux,"
        sSentencia = sSentencia & fIsNull() & "a.CodAux,'') CodAux,"
'fin 2015-12-07 libros inven y balance correcc
        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
        sSentencia = sSentencia & "    a.FeEDoc,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END), 0), 2) AS DebeSol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END), 0), 2) AS HaberSol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END), 0), 2) AS DebeDol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END), 0), 2) AS HaberDol, "
        sSentencia = sSentencia & fIsNull() & "a.CodTDc,'') CodTDc,"
        sSentencia = sSentencia & fIsNull() & "a.SerDoc,'') SerDoc,"
        sSentencia = sSentencia & fIsNull() & "a.NroDoc,'') NroDoc "
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 12, 15, 16, 18
        sSentencia = sSentencia & "," & fIsNull() & "a.codcta,'') codcta "
        End Select
'fin 2015-12-07 libros inven y balance correcc
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 18
        sSentencia = sSentencia & "," & fIsNull() & "a.gloite,'') gloite "
        End Select
'fin 2015-12-07 libros inven y balance correcc
        sSentencia = sSentencia & "From (COCpbDet a "
        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
        '2015-11-12/18 PLE rpt fmt electro archivo sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
        Select Case pnOpcion
        Case 6, 7
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
        Case 8
        sSentencia = sSentencia & " AND left(a.CodCta,2)='14'  "
        Case 9, 10
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='16'  OR left(a.CodCta,2)='17')  "
        Case 11
        sSentencia = sSentencia & " AND left(a.CodCta,2)='19'  "
        Case 12
        sSentencia = sSentencia & " AND left(a.CodCta,2)='41'  "
        Case 13, 14
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='42'  OR left(a.CodCta,2)='43')  "
        Case 15, 16
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='46'  OR left(a.CodCta,2)='47')  "
        Case 18
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='37'  OR left(a.CodCta,2)='49')  "
        Case Else
        sSentencia = sSentencia & " "
        End Select
        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
        
sSentencia = sSentencia & "AND (ROUND(ROUND(IFNULL((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END), 0), 2)"
sSentencia = sSentencia & "   - ROUND(IFNULL((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END), 0), 2), 2) <> 0.00"
sSentencia = sSentencia & "OR ROUND(ROUND(IFNULL((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END), 0), 2)"
sSentencia = sSentencia & " - ROUND(IFNULL((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END), 0), 2), 2) <> 0.00 )"
       
'        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux "
'        If ps_Plataforma = pSrvMySql Then
'            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
'        Else
'            sSentencia = sSentencia & " HAVING "
'            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
'            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
'        End If

'ini 2015-12-07 libros inven y balance correcc
''        sSentencia = sSentencia & "ORDER BY a.CodAux,mespvs,coddro,nrocpb,nroite"
        Select Case pnOpcion
        Case 12
        sSentencia = sSentencia & "ORDER BY a.CodCta,a.CodAux,mespvs,coddro,nrocpb,nroite"
        Case 15, 16, 18
        sSentencia = sSentencia & "ORDER BY a.CodAux,a.CodCta,mespvs,coddro,nrocpb,nroite"
        Case Else
        sSentencia = sSentencia & "ORDER BY a.CodAux,mespvs,coddro,nrocpb,nroite"
        End Select
'fin 2015-12-07 libros inven y balance correcc
        

'fin 2015-12-04 libros inven y balance correcc
        
'fin 2015-11-12/18 PLE rpt fmt electro archivo

'ini 2015-11-12/18 PLE rpt fmt electro archivo
'case=11 3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
'archivo agrupado y ordenado por RUC+tpodoc+serie+nrodoc
   
''''''ini 2015-12-07 libros inven y balance correcc
'''''   Case 11
'''''        Select Case pnOpcion
'''''        Case 11: sArchivo = "030600" & "00" & "1"
'''''        Case Else: sArchivo = "sin_archivo"
'''''        End Select
'''''        sSentencia = ""
'''''        sSentencia = sSentencia & "SELECT"
'''''        sSentencia = sSentencia & "    MIN(mespvs) mespvs,"
'''''        sSentencia = sSentencia & "    MIN(a.coddro) coddro,MIN(a.nrocpb) nrocpb,MIN(a.nroite) nroite,"
'''''        sSentencia = sSentencia & fIsNull() & "MAX(c.tpodci),'') tpodci,"
'''''        sSentencia = sSentencia & "    a.CodAux,"
'''''        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
'''''        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
'''''        '***
'''''        sSentencia = sSentencia & fIsNull() & "a.CodTDc,'') CodTDc," & fIsNull() & "a.SerDoc,'') SerDoc," & fIsNull() & "a.NroDoc,'') NroDoc,"
'''''        '***
'''''        sSentencia = sSentencia & "    MIN(a.FeEDoc) AS FeEDoc,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
'''''        sSentencia = sSentencia & "From (COCpbDet a "
'''''        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
'''''        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
'''''        Select Case pnOpcion
'''''        Case 11
'''''        sSentencia = sSentencia & " AND left(a.CodCta,2)='19'  "
'''''        Case Else
'''''        sSentencia = sSentencia & " "
'''''        End Select
'''''        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
'''''        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux, a.CodTDc, a.SerDoc, a.NroDoc "
'''''        If ps_Plataforma = pSrvMySql Then
'''''            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
'''''        Else
'''''            sSentencia = sSentencia & " HAVING "
'''''            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
'''''            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
'''''        End If
'''''        sSentencia = sSentencia & "ORDER BY a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
''''''fin 2015-12-07 libros inven y balance correcc
   Case 19   '3.17 LIBRO DE INVENTARIOS Y BALANCES - BALANCE DE COMPROBACIÓN (3)
'ini 2015-12-10 PLE rpt fmt electro archivo
    sArchivo = "031900" & "00" & "1"
    BalancedeComprobacion2
'''  With porstMRp
'''    If .State = adStateOpen Then .Close
    sSentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
    sSentencia = sSentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
    sSentencia = sSentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt, "
    sSentencia = sSentencia & "ROUND(SUM(cApeD), 2) AS cApeD, ROUND(SUM(cApeH), 2) AS cApeH "
    sSentencia = sSentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
    sSentencia = sSentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
    sSentencia = sSentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)+ ROUND(SUM(cApeD), 2)+ ROUND(SUM(cApeH), 2)) > 0 "
    sSentencia = sSentencia & "ORDER BY CodCta"
''    .Source = s_Sentencia
'''    .Open
'''  End With
    
'fin 2015-12-10 PLE rpt fmt electro archivo

'fin 2015-11-12/18 PLE rpt fmt electro archivo
   Case 23      ' libro diario
    ' Incializo variables
    sArchivo = "050100" & "00" & "1"
    sSentencia = fAE_05_1_lib_diario_sql(sMoneda)
    
   Case 24      ' libro diario simplificado
    ' Incializo variables
    sArchivo = "050200" & "00" & "1"
    sSentencia = fAE_05_2_lib_diario_simplificado_sql(sMoneda)
   Case 27      ' libro mayor
    ' Incializo variables
    sArchivo = "060100" & "00" & "1"
    sSentencia = fAE_06_1_lib_mayor_sql(sMoneda)
   Case 29, 30    ' registro  de compras
    sArchivo = "080100" & "00" & "1"
    sSentencia = fAE_08_1_reg_cpr_sql(sMoneda, 1)
   Case 31      ' registro de ventas
     sArchivo = "140100" & "00" & "1"
     sSentencia = fAE_14_1_reg_vta_sql(sMoneda)
  
'ini 2014-05-30 adicion 5.3 plan ctas
   Case 32      'Plan de cuentas
    sArchivo = "050300" & "00" & "1"
    sSentencia = fAE_05_3_lib_diario_deta_plan_cta_sql(sMoneda)
   Case Else: Exit Sub
  End Select
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "1", "2")
'ini 2016-02-02.19 correccion ple
  'sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "0" & sMoneda & "1" & ".txt"
  'segun teo cuado va a grabar el archivo, en la ventana de grabar archivo, sale
  'por ejemplo: LE2060087744620160100140100001011 y cuando graba se pone:LE2060087744620160100140100001111
  ' se hace las correcciones para los libros corregidos.
  ' registro  de compras / registro de ventas / Plan de cuentas
  '' libro diario / ' libro diario simplificado / libro mayor
  Select Case pnOpcion
  Case 23, 24, 27, 29, 30, 31, 32
    sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "1" & sMoneda & "1" & ".txt"
  Case Else
    sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "0" & sMoneda & "1" & ".txt"
  End Select
'fin 2016-02-02.19 correccion ple
  
  On Error GoTo CancelaDialogo
  cdlMain.DialogTitle = "Grabar Archivo Como"
  cdlMain.CancelError = True
  cdlMain.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
  cdlMain.FileName = sArchivo
  cdlMain.DefaultExt = ".txt"
  cdlMain.Filter = "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*"
  cdlMain.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  ChDir App.path
  If MsgBox("¿ Estás Seguro de Generar Registro Electrónico? ", vbQuestion + vbYesNo) = vbYes Then
    sArchivo = cdlMain.FileName
    sExpresion = cdlMain.FileTitle
    ppArchivoElectronico sArchivo, sExpresion, sSentencia
    MsgBox TEXT_8008, vbInformation
  End If
  ChDrive Left$(App.path, 1)
  ChDir App.path
  
'ini 2016-02-02.05 correccion ple
      Select Case pnOpcion
       Case 29, 30: ppRegistroElectronico2
      End Select
'fin 2016-02-02.05 correccion ple

End Sub
'ini 2016-02-02.05 correccion ple

Private Sub ppRegistroElectronico2()

  'ini rcs 2015-04-27 correccion version
        Dim arrDocume(4) As String
        arrDocume(0) = "d.AbvTDc"
        arrDocume(1) = "'-'"
        arrDocume(2) = "a.SerDoc"
        arrDocume(3) = "'-'"
        arrDocume(4) = "a.NroDoc"
  'ini rcs 2015-04-27 correccion version

  Dim sSentencia As String, sArchivo As String
  Dim sMoneda As String, sExpresion As String
  Dim s_MesIni As String, s_MesFin As String
  Dim sSalAntDeb As String, sSalAntHab As String
  Dim nSecuencia As Long
  If cboTpoMon.Text = "" Then MsgBox "Seleccionar Moneda", vbCritical, "Sistema Contable": Exit Sub
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "MN", "ME")
  xqmes = Left(cmbEjercicio.Text, 2)
  Select Case pnOpcion
   Case 29, 30    ' registro  de compras
   'segun liliana interfondos dice que "080200" & "00" & "2" cambia a "080200" & "00" & "1"
   '2016-02-02.15 correccion ple  sArchivo = "080200" & "00" & "2"
    sArchivo = "080200" & "00" & "1"
   sSentencia = fAE_08_2_reg_cpr_sql(sMoneda, 2)
   Case Else: Exit Sub
  End Select
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "1", "2")
  
  'ini 2016-02-02.19 correccion ple
  'sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "0" & sMoneda & "1" & ".txt"
  'segun teo cuado va a grabar el archivo, en la ventana de grabar archivo, sale
  'por ejemplo: LE2060087744620160100140100001011 y cuando graba se pone:LE2060087744620160100140100001111
  ' se hace las correcciones para los libros corregidos.
  ' registro  de compras
  Select Case pnOpcion
  Case 29, 30
    sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "1" & sMoneda & "1" & ".txt"
  Case Else
    sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "0" & sMoneda & "1" & ".txt"
  End Select
'fin 2016-02-02.19 correccion ple

  
  On Error GoTo CancelaDialogo
  cdlMain.DialogTitle = "Grabar Archivo Como"
  cdlMain.CancelError = True
  cdlMain.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
  cdlMain.FileName = sArchivo
  cdlMain.DefaultExt = ".txt"
  cdlMain.Filter = "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*"
  cdlMain.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  ChDir App.path
  If MsgBox("¿ Estás Seguro de Generar Registro Electrónico? ", vbQuestion + vbYesNo) = vbYes Then
    sArchivo = cdlMain.FileName
    sExpresion = cdlMain.FileTitle
    ppArchivoElectronico2 sArchivo, sExpresion, sSentencia
    MsgBox TEXT_8008, vbInformation
  End If
  ChDrive Left$(App.path, 1)
  ChDir App.path
End Sub

Private Sub ppArchivoElectronico2(ByVal sArchivo As String, ByVal sNombreArchivo As String, ByVal sSentencia As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  '2016-02-02.03 correccion ple Dim psRegistro As String, s_Caracter As String, s_Expresion As String
  Dim psRegistro As String, s_Expresion As String
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nRegistroAux As Long, nRegistroDeta As Long, nTamano As Integer
  Dim sAuxiliar As String, s_OldMessage As String
  Dim nSumatoriaTotal As Double
  
   Dim n_SdoDeb As Double, n_SdoHab As Double 'teo 2015-12-15 falta definir
 
  ' selecciono informacion de proceso
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
    nRegistros = .RecordCount
  End With

  ' Creo objeto de archivo
  If nRegistros > 0 Then
    s_Expresion = Left(sNombreArchivo, 30) & "1" & Mid(sNombreArchivo, 32)
    sArchivo = Replace(sArchivo, sNombreArchivo, s_Expresion)
'ini 2016-02-02.26 correccion ple archivo vacio=0
  Else
    s_Expresion = Left(sNombreArchivo, 30) & "0" & Mid(sNombreArchivo, 32)
    sArchivo = Replace(sArchivo, sNombreArchivo, s_Expresion)
'fin 2016-02-02.26 correccion ple archivo vacio=0
  End If
  Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
  Set potxtFileExp = pofsoFileExp.CreateTextFile(sArchivo, True)
  s_Caracter = "|"
  ' detalle de archivo
  Dim xxMesPvs As String
  xxMesPvs = Left(cmbEjercicio.Text, 2)
  'ini 2014-07-31 numero correlativo
  '2015-04-07 error desbordamiento Dim xNroCorr As Integer
  Dim xNroCorr As Long
  xNroCorr = 1
  'fin 2014-07-31 numero correlativo
  If Not (porstMRp.BOF And porstMRp.EOF) Then
    nRegistro = 0
    While Not porstMRp.EOF
      psRegistro = ""
      Select Case pnOpcion
       Case 29, 30: psRegistro = fAE_08_2_reg_cpr(xNroCorr)   ' registro compras
      End Select
      potxtFileExp.WriteLine psRegistro
      nRegistro = nRegistro + 1
      porstMRp.MoveNext
      xNroCorr = xNroCorr + 1 '2014-07-31 numero correlativo
    Wend
    ' Cierro objeto y saco de memoria
    potxtFileExp.Close
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  GoTo Finalizar
  
error:
Finalizar:
  ' Reinicializo los mensajes
  ' Coloco el puntero en normal
End Sub
'fin 2016-02-02.05 correccion ple

Private Sub ppRegistroElectronico_2015_01_22()

  'ini rcs 2015-04-27 correccion version
        Dim arrDocume(4) As String
        arrDocume(0) = "d.AbvTDc"
        arrDocume(1) = "'-'"
        arrDocume(2) = "a.SerDoc"
        arrDocume(3) = "'-'"
        arrDocume(4) = "a.NroDoc"
  'ini rcs 2015-04-27 correccion version

  Dim sSentencia As String, sArchivo As String
  Dim sMoneda As String, sExpresion As String
  Dim s_MesIni As String, s_MesFin As String
  Dim sSalAntDeb As String, sSalAntHab As String
  Dim nSecuencia As Long

  ' valido información
  If pnOpcion = 99 Then MsgBox "Seleccionar Libro o Registro", vbCritical, "Sistema Contable": Exit Sub
  If cboTpoMon.Text = "" Then MsgBox "Seleccionar Moneda", vbCritical, "Sistema Contable": Exit Sub
  
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "MN", "ME")
  xqmes = Left(cmbEjercicio.Text, 2)
  Select Case pnOpcion
   'Case 5: sSentencia = LibroCajaElectro(999): sArchivo = "050100" & "00" & "2": Exit Sub '2015-04-23 nvos.rpt teo 3.2 LIBRO DE INVENTARIOS Y BALANCES
   Case 5: sSentencia = LibroCajaElectro(999): sArchivo = "030100" & "00" & "2" '2015-04-23 nvos.rpt teo 3.2 LIBRO DE INVENTARIOS Y BALANCES
   
'ini 2015-11-12/18 PLE rpt fmt electro archivo

'ini 2015-12-04 libros inven y balance correcc
'   Case 6, 7, 8, 9, 10
'        '2015-11-12/18 PLE rpt fmt electro archivo sArchivo = "040100" & "00" & "1"
'        Select Case pnOpcion
'        Case 6, 7: sArchivo = "040100" & "00" & "1"
'        Case 8: sArchivo = "030400" & "00" & "1"
'        Case 9, 10: sArchivo = "030500" & "00" & "1"
'        Case Else: sArchivo = "sin_archivo"
'        End Select
'        sSentencia = ""
'        sSentencia = sSentencia & "SELECT"
'        sSentencia = sSentencia & "    MIN(mespvs) mespvs,"
'        sSentencia = sSentencia & "    MIN(a.coddro) coddro,MIN(a.nrocpb) nrocpb,MIN(a.nroite) nroite,"
'        sSentencia = sSentencia & fIsNull() & "MAX(c.tpodci),'') tpodci,"
'        sSentencia = sSentencia & "    a.CodAux,"
'        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
'        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
'        sSentencia = sSentencia & "    MIN(a.FeEDoc) AS FeEDoc,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol,"
'        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
'        sSentencia = sSentencia & "From (COCpbDet a "
'        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
'        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
'        '2015-11-12/18 PLE rpt fmt electro archivo sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
'        Select Case pnOpcion
'        Case 6, 7
'        sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
'        Case 8
'        sSentencia = sSentencia & " AND left(a.CodCta,2)='14'  "
'        Case 9, 10
'        sSentencia = sSentencia & " AND (left(a.CodCta,2)='16'  OR left(a.CodCta,2)='17')  "
'        Case Else
'        sSentencia = sSentencia & " "
'        End Select
'
'        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
'        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux "
'        'sSentencia = sSentencia & "Having (Round(DebeSol - HaberSol, 2) <> 0.00 Or Round(DebeDol - HaberDol, 2) <> 0.00) "
'        If ps_Plataforma = pSrvMySql Then
'            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
'        Else
'            sSentencia = sSentencia & " HAVING "
'            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
'            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
'        End If
'        sSentencia = sSentencia & "ORDER BY a.CodAux"
'
'case=6y7 3.3 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 12 CUENTAS POR COBRAR COMERCIALES – TERCEROS Y 13
'case=8 3.4 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS
'case=9y10 3.5 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS - TERCEROS O CUENTA 17
'case=11 3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
'case=12 3.11 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 41 REMUNERACIONES Y PARTICIPACIONES POR PAGAR (PCGE) (2)
'case=13y14 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=15y16 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=18 3.15  LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 37 ACTIVO DIFERIDO Y DE LA CUENTA 49 PASIVO DIFERIDO (PCGE)   (2)
'archivo agrupado y ordenado por RUC
   Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18
        '2015-11-12/18 PLE rpt fmt electro archivo sArchivo = "040100" & "00" & "1"
        Select Case pnOpcion
        Case 6, 7: sArchivo = "040100" & "00" & "1"
        Case 8: sArchivo = "030400" & "00" & "1"
        Case 9, 10: sArchivo = "030500" & "00" & "1"
        Case 11: sArchivo = "030600" & "00" & "1"
        Case 12: sArchivo = "031100" & "00" & "1"
        Case 13, 14: sArchivo = "041300" & "00" & "1"
        Case 15, 16: sArchivo = "041400" & "00" & "1"
        Case 18: sArchivo = "031500" & "00" & "1"
        Case Else: sArchivo = "sin_archivo"
        End Select
        sSentencia = ""
        sSentencia = sSentencia & "SELECT"
        sSentencia = sSentencia & "    a.mespvs,"
        sSentencia = sSentencia & "    a.coddro, a.nrocpb, a.nroite,"
        sSentencia = sSentencia & fIsNull() & "c.tpodci,'') tpodci,"
'ini 2015-12-07 libros inven y balance correcc
        'sSentencia = sSentencia & "    a.CodAux,"
        sSentencia = sSentencia & fIsNull() & "a.CodAux,'') CodAux,"
'fin 2015-12-07 libros inven y balance correcc
        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
        sSentencia = sSentencia & "    a.FeEDoc,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END), 0), 2) AS DebeSol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END), 0), 2) AS HaberSol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END), 0), 2) AS DebeDol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "(CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END), 0), 2) AS HaberDol, "
        sSentencia = sSentencia & fIsNull() & "a.CodTDc,'') CodTDc,"
        sSentencia = sSentencia & fIsNull() & "a.SerDoc,'') SerDoc,"
        sSentencia = sSentencia & fIsNull() & "a.NroDoc,'') NroDoc "
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 12, 15, 16, 18
        sSentencia = sSentencia & "," & fIsNull() & "a.codcta,'') codcta "
        End Select
'fin 2015-12-07 libros inven y balance correcc
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 18
        sSentencia = sSentencia & "," & fIsNull() & "a.gloite,'') gloite "
        End Select
'fin 2015-12-07 libros inven y balance correcc
        sSentencia = sSentencia & "From (COCpbDet a "
        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
        '2015-11-12/18 PLE rpt fmt electro archivo sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
        Select Case pnOpcion
        Case 6, 7
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
        Case 8
        sSentencia = sSentencia & " AND left(a.CodCta,2)='14'  "
        Case 9, 10
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='16'  OR left(a.CodCta,2)='17')  "
        Case 11
        sSentencia = sSentencia & " AND left(a.CodCta,2)='19'  "
        Case 12
        sSentencia = sSentencia & " AND left(a.CodCta,2)='41'  "
        Case 13, 14
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='42'  OR left(a.CodCta,2)='43')  "
        Case 15, 16
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='46'  OR left(a.CodCta,2)='47')  "
        Case 18
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='37'  OR left(a.CodCta,2)='49')  "
        Case Else
        sSentencia = sSentencia & " "
        End Select
        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
        
sSentencia = sSentencia & "AND (ROUND(ROUND(IFNULL((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END), 0), 2)"
sSentencia = sSentencia & "   - ROUND(IFNULL((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END), 0), 2), 2) <> 0.00"
sSentencia = sSentencia & "OR ROUND(ROUND(IFNULL((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END), 0), 2)"
sSentencia = sSentencia & " - ROUND(IFNULL((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END), 0), 2), 2) <> 0.00 )"
       
'        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux "
'        If ps_Plataforma = pSrvMySql Then
'            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
'        Else
'            sSentencia = sSentencia & " HAVING "
'            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
'            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
'        End If

'ini 2015-12-07 libros inven y balance correcc
''        sSentencia = sSentencia & "ORDER BY a.CodAux,mespvs,coddro,nrocpb,nroite"
        Select Case pnOpcion
        Case 12
        sSentencia = sSentencia & "ORDER BY a.CodCta,a.CodAux,mespvs,coddro,nrocpb,nroite"
        Case 15, 16, 18
        sSentencia = sSentencia & "ORDER BY a.CodAux,a.CodCta,mespvs,coddro,nrocpb,nroite"
        Case Else
        sSentencia = sSentencia & "ORDER BY a.CodAux,mespvs,coddro,nrocpb,nroite"
        End Select
'fin 2015-12-07 libros inven y balance correcc
        

'fin 2015-12-04 libros inven y balance correcc
        
'fin 2015-11-12/18 PLE rpt fmt electro archivo

'ini 2015-11-12/18 PLE rpt fmt electro archivo
'case=11 3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
'archivo agrupado y ordenado por RUC+tpodoc+serie+nrodoc
   
''''''ini 2015-12-07 libros inven y balance correcc
'''''   Case 11
'''''        Select Case pnOpcion
'''''        Case 11: sArchivo = "030600" & "00" & "1"
'''''        Case Else: sArchivo = "sin_archivo"
'''''        End Select
'''''        sSentencia = ""
'''''        sSentencia = sSentencia & "SELECT"
'''''        sSentencia = sSentencia & "    MIN(mespvs) mespvs,"
'''''        sSentencia = sSentencia & "    MIN(a.coddro) coddro,MIN(a.nrocpb) nrocpb,MIN(a.nroite) nroite,"
'''''        sSentencia = sSentencia & fIsNull() & "MAX(c.tpodci),'') tpodci,"
'''''        sSentencia = sSentencia & "    a.CodAux,"
'''''        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
'''''        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
'''''        '***
'''''        sSentencia = sSentencia & fIsNull() & "a.CodTDc,'') CodTDc," & fIsNull() & "a.SerDoc,'') SerDoc," & fIsNull() & "a.NroDoc,'') NroDoc,"
'''''        '***
'''''        sSentencia = sSentencia & "    MIN(a.FeEDoc) AS FeEDoc,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol,"
'''''        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
'''''        sSentencia = sSentencia & "From (COCpbDet a "
'''''        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
'''''        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
'''''        Select Case pnOpcion
'''''        Case 11
'''''        sSentencia = sSentencia & " AND left(a.CodCta,2)='19'  "
'''''        Case Else
'''''        sSentencia = sSentencia & " "
'''''        End Select
'''''        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
'''''        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux, a.CodTDc, a.SerDoc, a.NroDoc "
'''''        If ps_Plataforma = pSrvMySql Then
'''''            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
'''''        Else
'''''            sSentencia = sSentencia & " HAVING "
'''''            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
'''''            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
'''''        End If
'''''        sSentencia = sSentencia & "ORDER BY a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
''''''fin 2015-12-07 libros inven y balance correcc
   Case 19   '3.17 LIBRO DE INVENTARIOS Y BALANCES - BALANCE DE COMPROBACIÓN (3)
'ini 2015-12-10 PLE rpt fmt electro archivo
    sArchivo = "031900" & "00" & "1"
    BalancedeComprobacion2
'''  With porstMRp
'''    If .State = adStateOpen Then .Close
    sSentencia = "SELECT CodCta, DetCta, TpoSdo, TpoCta, "
    sSentencia = sSentencia & "ROUND(SUM(cSumaD), 2) AS cSumaD, ROUND(SUM(cSumaH), 2) AS cSumaH, "
    sSentencia = sSentencia & "ROUND(SUM(cSumaDt), 2) AS cSumaDt, ROUND(SUM(cSumaHt), 2) AS cSumaHt, "
    sSentencia = sSentencia & "ROUND(SUM(cApeD), 2) AS cApeD, ROUND(SUM(cApeH), 2) AS cApeH "
    sSentencia = sSentencia & "FROM " & ps_Prefijo & "trpRngBceCpb "
    sSentencia = sSentencia & "GROUP BY CodCta, DetCta, TpoSdo, TpoCta "
    sSentencia = sSentencia & "HAVING (ROUND(SUM(cSumaD), 2) + ROUND(SUM(cSumaH), 2) + ROUND(SUM(cSumaDt), 2) + ROUND(SUM(cSumaHt), 2)+ ROUND(SUM(cApeD), 2)+ ROUND(SUM(cApeH), 2)) > 0 "
    sSentencia = sSentencia & "ORDER BY CodCta"
''    .Source = s_Sentencia
'''    .Open
'''  End With
    
'fin 2015-12-10 PLE rpt fmt electro archivo

'fin 2015-11-12/18 PLE rpt fmt electro archivo
   Case 23      ' libro diario
    ' Incializo variables
    sArchivo = "050100" & "00" & "1"
    sSentencia = "SELECT det.CodDro, det.nrocpb, det.nroite, det.fehope, det.codtdc, det.serdoc, det.nrodoc, "
    sSentencia = sSentencia & "det.codcta, cta.detcta, det.codaux, aux.razaux, det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & "tdc.AbvTDc, " & Choose(gsIdioma, "dro.DetDro", "dro.DetDrox") & " AS DetDro, "
    sSentencia = sSentencia & "'" & sMoneda & "' as sMoneda, dro.codlib "
    sSentencia = sSentencia & ",det.mespvs " '2015-04-06 convierte un solo mes 12y00
    sSentencia = sSentencia & "FROM (((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
'ini 2015-04-01 convierte un solo mes 12y00
    'sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    If xqmes = "12" Then
        sSentencia = sSentencia & "AND det.mespvs >='" & xqmes & "'" & " AND det.mespvs <='13' "
    ElseIf xqmes = "01" Then
'ini 2015-04-20 corre 00 en ene dia
        sSentencia = sSentencia & "AND det.mespvs >='00'" & " AND det.mespvs <='" & xqmes & "' "
'fin 2015-04-20 corre 00 en ene dia
    Else
    sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    End If
'fin 2015-04-01 convierte un solo mes 12y00
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    '2015-04-01 convierte un solo mes 12y00 sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite, fehope"
    sSentencia = sSentencia & "ORDER BY mespvs, coddro, nrocpb, nroite, fehope"
   Case 24      ' libro diario simplificado
    ' Incializo variables
    sArchivo = "050200" & "00" & "1"
    sSentencia = "SELECT det.CodDro, det.nrocpb, det.nroite, det.fehope, det.codtdc, det.serdoc, det.nrodoc, "
    sSentencia = sSentencia & "det.codcta, cta.detcta, det.codaux, aux.razaux, det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & "tdc.AbvTDc, " & Choose(gsIdioma, "dro.DetDro", "dro.DetDrox") & " AS DetDro, "
    sSentencia = sSentencia & "'" & sMoneda & "' as sMoneda, dro.codlib "
    sSentencia = sSentencia & "FROM (((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite, fehope"
   Case 27      ' libro mayor
    ' Incializo variables
    sArchivo = "060100" & "00" & "1"
    s_MesIni = Left(cmbEjercicio.Text, 2)
    s_MesFin = Left(cmbEjercicio.Text, 2)
    ' movimientos
    '2015-04-07 error mespvs sSentencia = "SELECT  det.mespvs,det.codcta, det.coddro, det.nrocpb, det.nroite, det.fehope, "
    'sSentencia = "SELECT  CASE det.mespvs WHEN det.mespvs>='12' THEN det.mespvs ELSE " & s_MesIni & " END mespvs, "
    sSentencia = "SELECT " & IIf(s_MesIni = "12", "det.mespvs", "'" & s_MesIni & "'") & " as mespvs, "
    sSentencia = sSentencia & " det.codcta, det.coddro, det.nrocpb, det.nroite, det.fehope, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-' + det.serdoc + '-' + det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "det.codaux, aux.RazAux, det.refdoc, det.tpodoc AS medpago, "
    sSentencia = sSentencia & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, " & Choose(gsIdioma, "dro.detdro", "dro.detdrox") & " AS detdro, tdc.abvtdc "
    'sSentencia = sSentencia & ",det.mespvs " '2015-04-06 convierte un solo mes 12y00
    sSentencia = sSentencia & "FROM ((((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
'ini 2015-04-01 convierte un solo mes 12y00
    'sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & s_MesFin & "' "
    If s_MesIni = "12" Then
    sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & "13" & "' "
    Else
    sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & s_MesFin & "' "
    End If
'fin 2015-04-01 convierte un solo mes 12y00
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    ' Saldo anterior
    '2015-04-20 solo sdo ante de ene a dic If s_MesIni <> "00" Then
    'If (s_MesIni <> "00") Or Not (s_MesIni >= "02" And s_MesIni <= "12") Then
    If s_MesIni = "01" Or s_MesIni = "13" Then
      sSalAntDeb = "ROUND(("
      sSalAntHab = "ROUND(("
      For nSecuencia = 0 To (Val(s_MesIni) - 1)
        sSalAntDeb = sSalAntDeb & "acu.acud" & Format(nSecuencia, "00") & "_" & sMoneda & IIf(nSecuencia = (Val(s_MesIni) - 1), ")", "+")
        sSalAntHab = sSalAntHab & "acu.acuh" & Format(nSecuencia, "00") & "_" & sMoneda & IIf(nSecuencia = (Val(s_MesIni) - 1), ")", "+")
      Next nSecuencia
      sSalAntDeb = sSalAntDeb & ", 2)"
      sSalAntHab = sSalAntHab & ", 2)"
      '2015-04-20 corre 00 en ene dia
      sSentencia = sSentencia & "UNION "
      '2015-04-07 error mespvs sSentencia = sSentencia & "SELECT '00' AS mespvs, cta.codcta, Null, Null, Null, Null, Null, Null, Null, Null, Null, "
      'sSentencia = sSentencia & "SELECT " & IIf(s_MesIni = "12", "'" & s_MesIni & "'", "'00'") & " AS mespvs, "
      sSentencia = sSentencia & "SELECT " & "'" & s_MesIni & "'" & " AS mespvs, "
      sSentencia = sSentencia & " cta.codcta, Null, Null, Null, Null, Null, Null, Null, Null, Null, "
      sSentencia = sSentencia & "'" & Choose(gsIdioma, "SALDO ANTERIOR", "PREVIOUS BALANCE") & "' AS gloite, "
      sSentencia = sSentencia & sSalAntDeb & " AS nDebe, "
      sSentencia = sSentencia & sSalAntHab & " AS nHaber, "
      sSentencia = sSentencia & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, Null, Null "
      sSentencia = sSentencia & "FROM cocta cta "
      sSentencia = sSentencia & "LEFT JOIN coctaacu acu ON acu.codemp=cta.codemp AND acu.pdoano=cta.pdoano AND acu.codcta=cta.codcta "
      sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND cta.tpocta='" & TPOCTA_TRA & "' "
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & "HAVING (ROUND((nDebe-nHaber), 2)<>0.00) "
      Else
        sSentencia = sSentencia & "AND (ROUND((" & sSalAntDeb & "-" & sSalAntHab & "), 2)<>0.00) "
      End If
      '2015-04-20 corre 00 en ene dia
      
    End If
    sSentencia = sSentencia & "ORDER BY codcta, mespvs, coddro, nrocpb, nroite"
   Case 29, 30    ' registro  de compras
    sArchivo = "080100" & "00" & "1"
    sSentencia = "SELECT concat(com.coddro,com.nrocpb) as nrocpb, date_format(com.feedoc,'%d/%m/%Y') as feedoc, date_format(com.fevdoc,'%d/%m/%Y') as fevdoc, com.codtdc as codtdc, "
    sSentencia = sSentencia & "IFNULL(case com.codtdc when '50' then com.codaduana when '52' then com.codaduana when '53' then com.codaduana else com.serdoc end, '-') as serdoc, "
    sSentencia = sSentencia & "IFNULL(com.annodua, '0') as anno, IFNULL(CASE com.codtdc when '50' then com.nrodua when '52' then com.nrodua when '53' then com.nrodua else com.nrodoc END, '') as nrodoc, "
    sSentencia = sSentencia & "right(aux.tpodci,1) as tpodci, com.codaux as codaux, aux.razaux as razaux, "
    sSentencia = sSentencia & "com.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp1, "
    sSentencia = sSentencia & "com.impigv_ogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp2, "
    sSentencia = sSentencia & "com.impogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp3, "
    sSentencia = sSentencia & "com.impigv_ogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp4, "
    sSentencia = sSentencia & "com.impong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp5, "
    sSentencia = sSentencia & "com.impigv_ong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp6, "
    sSentencia = sSentencia & "com.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp7, "
    sSentencia = sSentencia & "com.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp8, "
    sSentencia = sSentencia & "com.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp9, "
    sSentencia = sSentencia & "com.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp10, "
    sSentencia = sSentencia & "CASE WHEN com.codtdc_ref='91' THEN Concat(com.serdoc_ref, '-', com.nrodoc_ref) ELSE Null END as v1, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "com.nrocdt, com.fehcdt, com.imptcb, com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
    sSentencia = sSentencia & "com.nrocdt, com.fehcdt,"
    sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN com.imptcb ELSE 0 END imptcb,"
    sSentencia = sSentencia & "com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
'fin 2015-09-16 t.cam solo ME
    'ini 2014-05-25
    sSentencia = sSentencia & ",codaduana,annodua,nrodua "
    'fin 2014-05-25
    sSentencia = sSentencia & ", date_format(com.fehope,'%d/%m/%Y') as fehope " '215-08-10 adicionado por teo
    sSentencia = sSentencia & "FROM cocprdoc com "
    sSentencia = sSentencia & "INNER JOIN tgaux aux ON com.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc ON com.codemp=tdc.codemp and com.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE com.codemp='" & gsCodEmp & "' AND com.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    sSentencia = sSentencia & "ORDER BY 1"
   Case 31      ' registro de ventas
    sArchivo = "140100" & "00" & "1"
    sSentencia = "SELECT concat(vta.coddro,vta.nrocpb) as nrocpb, date_format(vta.feedoc,'%d/%m/%Y') as feedoc, date_format(vta.fevdoc,'%d/%m/%Y') as fevdoc, vta.codtdc as codtdc, vta.serdoc as serdoc, "
    sSentencia = sSentencia & "vta.nrodoc, vta.nrodoc_fin, aux.tpodci as tpodci, (CASE WHEN tpodci='01' THEN RIGHT(aux.rucaux, 8) ELSE aux.rucaux END) as codaux, trim(left(aux.razaux,60)) as razaux, "
    sSentencia = sSentencia & "ROUND(vta.impexp_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impexp, "
    sSentencia = sSentencia & "ROUND(vta.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impogr, "
    sSentencia = sSentencia & "(CASE WHEN vta.categoriadoc<>'" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impexo, "
    sSentencia = sSentencia & "(CASE WHEN vta.categoriadoc='" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impina, "
    sSentencia = sSentencia & "ROUND(vta.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impisc, "
    sSentencia = sSentencia & "ROUND(vta.impigv_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impigv, "
    sSentencia = sSentencia & "ROUND(vta.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impoim, "
    sSentencia = sSentencia & "ROUND(vta.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as imptot, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "vta.imptcb as imptcb, "
    sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN vta.imptcb ELSE 0 END imptcb, "
'fin 2015-09-16 t.cam solo ME
    sSentencia = sSentencia & "date_format(feedoc_ref,'%d/%m/%Y') as d1, codtdc_ref as d2, serdoc_ref as d3, nrodoc_ref as d4, '" & sMoneda & "' as Moneda "
    'ini 2014-05-25
    sSentencia = sSentencia & ",impfob_mn "
    'fin 2014-05-25
    sSentencia = sSentencia & "FROM covtadoc vta "
    sSentencia = sSentencia & "INNER JOIN tgaux aux on vta.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc on vta.codemp=tdc.codemp and vta.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE vta.codemp='" & gsCodEmp & "' and vta.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    sSentencia = sSentencia & "ORDER BY vta.codtdc,vta.serdoc,vta.nrodoc"
    
'ini 2014-05-30 adicion 5.3 plan ctas
   Case 32      'Plan de cuentas
    sArchivo = "050300" & "00" & "1"
    'gsNivCta
    Dim xxniv1 As String
    Dim xxnivf As String
    xxniv1 = Left(gsNivCta, 1)
    xxnivf = Right(gsNivCta, 1)
    sSentencia = "SELECT"
    sSentencia = sSentencia & "      a.codcta , a.detcta, b.codcta codcta2, b.detcta detcta2 "
    sSentencia = sSentencia & "FROM cocta a "
    sSentencia = sSentencia & "LEFT JOIN cocta b "
    sSentencia = sSentencia & "    ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.codcta," & xxniv1 & ")=LEFT(b.codcta," & xxniv1 & ") "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "    AND a.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "    AND LENGTH(a.codcta)=" & xxnivf
    sSentencia = sSentencia & "    AND  LENGTH(b.codcta)=" & xxniv1 & " "
    sSentencia = sSentencia & "ORDER BY codcta"
'fin 2014-05-30 adicion 5.3 plan ctas

   Case Else: Exit Sub
  End Select
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "1", "2")
  sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "0" & sMoneda & "1" & ".txt"
  
  On Error GoTo CancelaDialogo
  cdlMain.DialogTitle = "Grabar Archivo Como"
  cdlMain.CancelError = True
  cdlMain.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
  cdlMain.FileName = sArchivo
  cdlMain.DefaultExt = ".txt"
  cdlMain.Filter = "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*"
  cdlMain.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  ChDir App.path
  If MsgBox("¿ Estás Seguro de Generar Registro Electrónico? ", vbQuestion + vbYesNo) = vbYes Then
    sArchivo = cdlMain.FileName
    sExpresion = cdlMain.FileTitle
    ppArchivoElectronico sArchivo, sExpresion, sSentencia
    MsgBox TEXT_8008, vbInformation
  End If
  ChDrive Left$(App.path, 1)
  ChDir App.path

End Sub

Private Sub ppRegistroElectronico_2015_11_18()

  'ini rcs 2015-04-27 correccion version
        Dim arrDocume(4) As String
        arrDocume(0) = "d.AbvTDc"
        arrDocume(1) = "'-'"
        arrDocume(2) = "a.SerDoc"
        arrDocume(3) = "'-'"
        arrDocume(4) = "a.NroDoc"
  'ini rcs 2015-04-27 correccion version

  Dim sSentencia As String, sArchivo As String
  Dim sMoneda As String, sExpresion As String
  Dim s_MesIni As String, s_MesFin As String
  Dim sSalAntDeb As String, sSalAntHab As String
  Dim nSecuencia As Long

  ' valido información
  If pnOpcion = 99 Then MsgBox "Seleccionar Libro o Registro", vbCritical, "Sistema Contable": Exit Sub
  If cboTpoMon.Text = "" Then MsgBox "Seleccionar Moneda", vbCritical, "Sistema Contable": Exit Sub
  
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "MN", "ME")
  xqmes = Left(cmbEjercicio.Text, 2)
  Select Case pnOpcion
   'Case 5: sSentencia = LibroCajaElectro(999): sArchivo = "050100" & "00" & "2": Exit Sub '2015-04-23 nvos.rpt teo 3.2 LIBRO DE INVENTARIOS Y BALANCES
   Case 5: sSentencia = LibroCajaElectro(999): sArchivo = "030100" & "00" & "2" '2015-04-23 nvos.rpt teo 3.2 LIBRO DE INVENTARIOS Y BALANCES
   
'ini 2015-11-12/18 PLE rpt fmt electro archivo
   Case 6, 7, 8
        '2015-11-12/18 PLE rpt fmt electro archivo sArchivo = "040100" & "00" & "1"
        Select Case pnOpcion
        Case 6, 7
        sArchivo = "040100" & "00" & "1"
        Case 8
        Case Else
        sArchivo = "sin_archivo"
        End Select
        sSentencia = ""
        sSentencia = sSentencia & "SELECT"
        sSentencia = sSentencia & "    MIN(mespvs) mespvs,"
        sSentencia = sSentencia & "    MIN(a.coddro) coddro,MIN(a.nrocpb) nrocpb,MIN(a.nroite) nroite,"
        sSentencia = sSentencia & fIsNull() & "MAX(c.tpodci),'') tpodci,"
        sSentencia = sSentencia & "    a.CodAux,"
        sSentencia = sSentencia & fIsNull() & "c.RucAux,'') RucAux,"
        sSentencia = sSentencia & fIsNull() & "c.RazAux,'') RazAux,"
        sSentencia = sSentencia & "    MIN(a.FeEDoc) AS FeEDoc,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol,"
        sSentencia = sSentencia & "    ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
        sSentencia = sSentencia & "From (COCpbDet a "
        sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
        
        '2015-11-12/18 PLE rpt fmt electro archivo sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
        Select Case pnOpcion
        Case 6, 7
        sSentencia = sSentencia & " AND (left(a.CodCta,2)='12'  OR left(a.CodCta,2)='13')  "
        Case 8
        sSentencia = sSentencia & " AND left(a.CodCta,2)='14'  "
        Case Else
        sSentencia = sSentencia & " "
        End Select
        
        sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2) & " "
        sSentencia = sSentencia & "GROUP BY  a.CodAux, c.RucAux, c.RazAux "
        'sSentencia = sSentencia & "Having (Round(DebeSol - HaberSol, 2) <> 0.00 Or Round(DebeDol - HaberDol, 2) <> 0.00) "
        If ps_Plataforma = pSrvMySql Then
            sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
        Else
            sSentencia = sSentencia & " HAVING "
            sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
            sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
        End If
        sSentencia = sSentencia & "ORDER BY a.CodAux"
'fin 2015-11-12/18 PLE rpt fmt electro archivo
    
   '2015-11-12 PLE rpt fmt electro archivo  Case 6
   '2015-11-12/18 PLE rpt fmt electro archivo  Case 6, 7
   Case 9996, 9997 'este es el original al 2015-11-18
    sArchivo = "040100" & "00" & "1"
    '2015-04-29 rpt 3.3 sSentencia = "SELECT  mespvs,a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite ,MAX(a.refdoc) refdoc ,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
     sSentencia = "SELECT  mespvs"
     sSentencia = sSentencia & ",a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, " & fIsNull() & "MAX(c.tpodci),'') tpodci, " & fIsNull() & "c.RucAux,'') RucAux, " & fIsNull() & "c.RazAux,'') RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite ,MAX(a.refdoc) refdoc ,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
   
    sSentencia = sSentencia & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sSentencia = sSentencia & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sSentencia = sSentencia & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sSentencia = sSentencia & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
'ini 2015-11-12 PLE rpt fmt electro archivo
    'sSentencia = sSentencia & " AND left(a.CodCta,2)='12' "
    sSentencia = sSentencia & " AND (left(a.CodCta,2)='12' "
    sSentencia = sSentencia & " OR left(a.CodCta,2)='13') "
'fin 2015-11-12 PLE rpt fmt electro archivo
    'sSentencia = sSentencia & " AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 "
    sSentencia = sSentencia & " and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sSentencia8 2015-03-23 correccion version sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sSentencia = sSentencia & " HAVING "
        sSentencia = sSentencia & "(ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
        sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sSentencia = sSentencia & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
'ini 2015-11-12/13 PLE rpt fmt electro archivo
   Case 8      '3.4 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS
    sArchivo = "030400" & "00" & "1"
    'sSentencia8 2015-03-23 sSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci), c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sSentencia = sSentencia & ",MAX(mespvs) mespvs " & "," & fIsNull() & "MAX(c.tpodci),'') tpodci "
    sSentencia = sSentencia & ", " & fIsNull() & "c.RucAux,'') RucAux, " & fIsNull() & "c.RazAux,'') RazAux "
    sSentencia = sSentencia & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sSentencia = sSentencia & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sSentencia = sSentencia & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sSentencia = sSentencia & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sSentencia = sSentencia & " AND left(a.CodCta,2)='14' AND IFNULL(a.CodAux, '') <>''"
    'sSentencia = sSentencia & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " AND left(a.CodCta,2)='14' "
    sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    '2015-11-12/13 PLE rpt fmt electro archivo  If ps_Plataforma = pSrvMysSentencia Then
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sSentencia = sSentencia & " HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00"
        sSentencia = sSentencia & "      OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
     sSentencia = sSentencia & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
   
'fin 2015-11-12/13 PLE rpt fmt electro archivo
'ini 2015-11-12/16 PLE rpt fmt electro archivo
   Case 9, 10 '3.5 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS - TERCEROS O CUENTA 17
   
'9ini ***********************************
    sArchivo = "030500" & "00" & "1"
    'sSentencia8 2015-03-23 sSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    
    sSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, " & fIsNull() & "MAX(c.tpodci),'') tpodci, " & fIsNull() & "c.RucAux,'') RucAux, " & fIsNull() & "c.RazAux,'') RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    '2015-11-12/13 PLEsSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sSentencia = sSentencia & ",MAX(a.mespvs) mespvs "
    sSentencia = sSentencia & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sSentencia = sSentencia & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
    sSentencia = sSentencia & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc) "
    sSentencia = sSentencia & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    ''2015-11-16 monetaneamente sSentencia = sSentencia & " AND left(a.CodCta,2)='16'  AND IFNULL(a.CodAux, '') <>''"
    sSentencia = sSentencia & " AND left(a.CodCta,2)='16'  "
    sSentencia = sSentencia & " AND mespvs <= " & Left(cmbEjercicio.Text, 2)
    'sSentencia = sSentencia & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    
      If ps_Plataforma = pSrvMySql Then
    '2015-11-12/14 PLE rpt fmt electro archivo   If ps_Plataforma = pSrvMysSentencia Then
        sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sSentencia = sSentencia & " HAVING "
        sSentencia = sSentencia & "  (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 "
        sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    '2015-11-12/14 PLE rpt fmt electro archivo  sSentencia = sSentencia & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
'9fin*************************************
        sSentencia = sSentencia & " UNION "
'10ini ***********************************
    'sSentencia8 2015-03-23 sSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    sSentencia = sSentencia & "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & "  AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,MAX(a.gloite) gloite,MAX(a.refdoc) refdoc,MAX(a.coddro) coddro,MAX(a.nrocpb) nrocpb "
    sSentencia = sSentencia & ",MAX(a.mespvs) mespvs "
    sSentencia = sSentencia & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sSentencia = sSentencia & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sSentencia = sSentencia & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sSentencia = sSentencia & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " AND left(a.CodCta,2)='17'  AND IFNULL(a.CodAux, '') <>''"
    sSentencia = sSentencia & " AND left(a.CodCta,2)='17'  AND " & fIsNull() & "a.CodAux, '') <>''"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    
    ''2015-11-16 monetaneamentesSentencia = sSentencia & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
      If ps_Plataforma = pSrvMySql Then
'2015-11-12/14 PLE rpt fmt electro archivo     If ps_Plataforma = pSrvMysSentencia Then
        sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sSentencia = sSentencia & " HAVING "
        sSentencia = sSentencia & " (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sSentencia = sSentencia & "  ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    
    End If
    '2015-11-12/14 PLE rpt fmt electro archivo sSentencia = sSentencia & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    sSentencia = sSentencia & " ORDER BY CodCta, CodAux, CodTDc, SerDoc, NroDoc"
'10ini ***********************************
   Case 11 '3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
   
    sArchivo = "030600" & "00" & "1"
    'sql8 2015-03-23 sql = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(IFNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, c.tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda,a.gloite,a.refdoc,a.coddro,a.nrocpb "
    '2015-11-12/16 PLE rpt fmt electro archivo sSentencia = "SELECT  a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda, MAX(a.gloite)gloite, MAX(a.refdoc) refdoc, MAX(a.coddro) coddro, MAX(a.nrocpb) nrocpb "
    sSentencia = "SELECT  a.CodCta, a.CodAux, " & fIsNull() & "a.CodTDc, '') CodTDc, " & fIsNull() & "a.SerDoc, '') SerDoc , " & fIsNull() & "a.NroDoc, '') NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, " & fConCat(arrDocume) & " AS cDocume, (CASE b.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, ROUND(" & fIsNull() & "SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, MAX(c.tpodci) tpodci, c.RucAux, c.RazAux, b.DetCta AS DetCta,'" & sMoneda & "' as Moneda, MAX(a.gloite)gloite, MAX(a.refdoc) refdoc, MIN(a.coddro) coddro, MIN(a.nrocpb) nrocpb "
    sSentencia = sSentencia & ", MAX(a.NroIte) NroIte , MIN(a.mespvs) mespvs"
    sSentencia = sSentencia & " From (((COCpbDet a LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta)"
    sSentencia = sSentencia & " LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux)"
    sSentencia = sSentencia & " LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc)"
    sSentencia = sSentencia & " WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "'"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " AND left(a.CodCta,2)='19' AND IFNULL(a.CodAux, '') <>''"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' AND IFNULL(a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " AND left(a.CodCta,2)='19' AND " & fIsNull() & "a.CodAux, '') <>''"
    '***** temporalmente 2015-11-12/16 PLE rpt fmt electro archivo sSentencia = sSentencia & " AND " & fIsNull() & "a.CodTDc, '') <>'' AND " & fIsNull() & "a.SerDoc, '') <>'' AND " & fIsNull() & "a.NroDoc, '') <>'' AND b.inddoc=1 and mespvs <= " & Left(cmbEjercicio.Text, 2)
    sSentencia = sSentencia & " GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, b.DetCta, b.tpomon"
    'sSentencia8 2015-03-23 sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    If ps_Plataforma = pSrvMySql Then
    '2015-11-12/16 PLE rpt fmt electro archivo If ps_Plataforma = pSrvMysSentencia Then
         sSentencia = sSentencia & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    Else
        sSentencia = sSentencia & " HAVING "
        sSentencia = sSentencia & " (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 OR"
        sSentencia = sSentencia & "  ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END)), 0), 2) - ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00)"
    End If
    sSentencia = sSentencia & " ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"


'fin 2015-11-12/16 PLE rpt fmt electro archivo
   Case 23      ' libro diario
    ' Incializo variables
    sArchivo = "050100" & "00" & "1"
    sSentencia = "SELECT det.CodDro, det.nrocpb, det.nroite, det.fehope, det.codtdc, det.serdoc, det.nrodoc, "
    sSentencia = sSentencia & "det.codcta, cta.detcta, det.codaux, aux.razaux, det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & "tdc.AbvTDc, " & Choose(gsIdioma, "dro.DetDro", "dro.DetDrox") & " AS DetDro, "
    sSentencia = sSentencia & "'" & sMoneda & "' as sMoneda, dro.codlib "
    sSentencia = sSentencia & ",det.mespvs " '2015-04-06 convierte un solo mes 12y00
    sSentencia = sSentencia & "FROM (((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
'ini 2015-04-01 convierte un solo mes 12y00
    'sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    If xqmes = "12" Then
        sSentencia = sSentencia & "AND det.mespvs >='" & xqmes & "'" & " AND det.mespvs <='13' "
    ElseIf xqmes = "01" Then
'ini 2015-04-20 corre 00 en ene dia
        sSentencia = sSentencia & "AND det.mespvs >='00'" & " AND det.mespvs <='" & xqmes & "' "
'fin 2015-04-20 corre 00 en ene dia
    Else
    sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    End If
'fin 2015-04-01 convierte un solo mes 12y00
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    '2015-04-01 convierte un solo mes 12y00 sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite, fehope"
    sSentencia = sSentencia & "ORDER BY mespvs, coddro, nrocpb, nroite, fehope"
   Case 24      ' libro diario simplificado
    ' Incializo variables
    sArchivo = "050200" & "00" & "1"
    sSentencia = "SELECT det.CodDro, det.nrocpb, det.nroite, det.fehope, det.codtdc, det.serdoc, det.nrodoc, "
    sSentencia = sSentencia & "det.codcta, cta.detcta, det.codaux, aux.razaux, det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & "tdc.AbvTDc, " & Choose(gsIdioma, "dro.DetDro", "dro.DetDrox") & " AS DetDro, "
    sSentencia = sSentencia & "'" & sMoneda & "' as sMoneda, dro.codlib "
    sSentencia = sSentencia & "FROM (((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND det.mespvs ='" & xqmes & "' "
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite, fehope"
   Case 27      ' libro mayor
    ' Incializo variables
    sArchivo = "060100" & "00" & "1"
    s_MesIni = Left(cmbEjercicio.Text, 2)
    s_MesFin = Left(cmbEjercicio.Text, 2)
    ' movimientos
    '2015-04-07 error mespvs sSentencia = "SELECT  det.mespvs,det.codcta, det.coddro, det.nrocpb, det.nroite, det.fehope, "
    'sSentencia = "SELECT  CASE det.mespvs WHEN det.mespvs>='12' THEN det.mespvs ELSE " & s_MesIni & " END mespvs, "
    sSentencia = "SELECT " & IIf(s_MesIni = "12", "det.mespvs", "'" & s_MesIni & "'") & " as mespvs, "
    sSentencia = sSentencia & " det.codcta, det.coddro, det.nrocpb, det.nroite, det.fehope, "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc+'-' + det.serdoc + '-' + det.nrodoc)") & " AS cDocume, "
    sSentencia = sSentencia & "det.codaux, aux.RazAux, det.refdoc, det.tpodoc AS medpago, "
    sSentencia = sSentencia & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nDebe, "
    sSentencia = sSentencia & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END) AS nHaber, "
    sSentencia = sSentencia & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, " & Choose(gsIdioma, "dro.detdro", "dro.detdrox") & " AS detdro, tdc.abvtdc "
    'sSentencia = sSentencia & ",det.mespvs " '2015-04-06 convierte un solo mes 12y00
    sSentencia = sSentencia & "FROM ((((cocpbdet det "
    sSentencia = sSentencia & "LEFT JOIN tgaux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux) "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta) "
    sSentencia = sSentencia & "LEFT JOIN codro dro ON dro.codemp=det.codemp AND dro.pdoano=det.pdoano AND dro.coddro=det.coddro) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
'ini 2015-04-01 convierte un solo mes 12y00
    'sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & s_MesFin & "' "
    If s_MesIni = "12" Then
    sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & "13" & "' "
    Else
    sSentencia = sSentencia & "AND det.MesPvs>='" & s_MesIni & "' AND det.MesPvs<='" & s_MesFin & "' "
    End If
'fin 2015-04-01 convierte un solo mes 12y00
    sSentencia = sSentencia & "AND det.imp" & sMoneda & "<>0.00 "
    ' Saldo anterior
    '2015-04-20 solo sdo ante de ene a dic If s_MesIni <> "00" Then
    'If (s_MesIni <> "00") Or Not (s_MesIni >= "02" And s_MesIni <= "12") Then
    If s_MesIni = "01" Or s_MesIni = "13" Then
      sSalAntDeb = "ROUND(("
      sSalAntHab = "ROUND(("
      For nSecuencia = 0 To (Val(s_MesIni) - 1)
        sSalAntDeb = sSalAntDeb & "acu.acud" & Format(nSecuencia, "00") & "_" & sMoneda & IIf(nSecuencia = (Val(s_MesIni) - 1), ")", "+")
        sSalAntHab = sSalAntHab & "acu.acuh" & Format(nSecuencia, "00") & "_" & sMoneda & IIf(nSecuencia = (Val(s_MesIni) - 1), ")", "+")
      Next nSecuencia
      sSalAntDeb = sSalAntDeb & ", 2)"
      sSalAntHab = sSalAntHab & ", 2)"
      '2015-04-20 corre 00 en ene dia
      sSentencia = sSentencia & "UNION "
      '2015-04-07 error mespvs sSentencia = sSentencia & "SELECT '00' AS mespvs, cta.codcta, Null, Null, Null, Null, Null, Null, Null, Null, Null, "
      'sSentencia = sSentencia & "SELECT " & IIf(s_MesIni = "12", "'" & s_MesIni & "'", "'00'") & " AS mespvs, "
      sSentencia = sSentencia & "SELECT " & "'" & s_MesIni & "'" & " AS mespvs, "
      sSentencia = sSentencia & " cta.codcta, Null, Null, Null, Null, Null, Null, Null, Null, Null, "
      sSentencia = sSentencia & "'" & Choose(gsIdioma, "SALDO ANTERIOR", "PREVIOUS BALANCE") & "' AS gloite, "
      sSentencia = sSentencia & sSalAntDeb & " AS nDebe, "
      sSentencia = sSentencia & sSalAntHab & " AS nHaber, "
      sSentencia = sSentencia & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, Null, Null "
      sSentencia = sSentencia & "FROM cocta cta "
      sSentencia = sSentencia & "LEFT JOIN coctaacu acu ON acu.codemp=cta.codemp AND acu.pdoano=cta.pdoano AND acu.codcta=cta.codcta "
      sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND cta.tpocta='" & TPOCTA_TRA & "' "
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & "HAVING (ROUND((nDebe-nHaber), 2)<>0.00) "
      Else
        sSentencia = sSentencia & "AND (ROUND((" & sSalAntDeb & "-" & sSalAntHab & "), 2)<>0.00) "
      End If
      '2015-04-20 corre 00 en ene dia
      
    End If
    sSentencia = sSentencia & "ORDER BY codcta, mespvs, coddro, nrocpb, nroite"
   Case 29, 30    ' registro  de compras
    sArchivo = "080100" & "00" & "1"
    sSentencia = "SELECT concat(com.coddro,com.nrocpb) as nrocpb, date_format(com.feedoc,'%d/%m/%Y') as feedoc, date_format(com.fevdoc,'%d/%m/%Y') as fevdoc, com.codtdc as codtdc, "
    sSentencia = sSentencia & "IFNULL(case com.codtdc when '50' then com.codaduana when '52' then com.codaduana when '53' then com.codaduana else com.serdoc end, '-') as serdoc, "
    sSentencia = sSentencia & "IFNULL(com.annodua, '0') as anno, IFNULL(CASE com.codtdc when '50' then com.nrodua when '52' then com.nrodua when '53' then com.nrodua else com.nrodoc END, '') as nrodoc, "
    sSentencia = sSentencia & "right(aux.tpodci,1) as tpodci, com.codaux as codaux, aux.razaux as razaux, "
    sSentencia = sSentencia & "com.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp1, "
    sSentencia = sSentencia & "com.impigv_ogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp2, "
    sSentencia = sSentencia & "com.impogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp3, "
    sSentencia = sSentencia & "com.impigv_ogn_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp4, "
    sSentencia = sSentencia & "com.impong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp5, "
    sSentencia = sSentencia & "com.impigv_ong_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp6, "
    sSentencia = sSentencia & "com.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp7, "
    sSentencia = sSentencia & "com.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp8, "
    sSentencia = sSentencia & "com.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp9, "
    sSentencia = sSentencia & "com.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end) as imp10, "
    sSentencia = sSentencia & "CASE WHEN com.codtdc_ref='91' THEN Concat(com.serdoc_ref, '-', com.nrodoc_ref) ELSE Null END as v1, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "com.nrocdt, com.fehcdt, com.imptcb, com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
    sSentencia = sSentencia & "com.nrocdt, com.fehcdt,"
    sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN com.imptcb ELSE 0 END imptcb,"
    sSentencia = sSentencia & "com.feedoc_ref, com.codtdc_ref, com.serdoc_ref, com.nrodoc_ref, com.indreten, '" & sMoneda & "' as Moneda "
'fin 2015-09-16 t.cam solo ME
    'ini 2014-05-25
    sSentencia = sSentencia & ",codaduana,annodua,nrodua "
    'fin 2014-05-25
    sSentencia = sSentencia & ", date_format(com.fehope,'%d/%m/%Y') as fehope " '215-08-10 adicionado por teo
    sSentencia = sSentencia & "FROM cocprdoc com "
    sSentencia = sSentencia & "INNER JOIN tgaux aux ON com.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc ON com.codemp=tdc.codemp and com.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE com.codemp='" & gsCodEmp & "' AND com.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    sSentencia = sSentencia & "ORDER BY 1"
   Case 31      ' registro de ventas
    sArchivo = "140100" & "00" & "1"
    sSentencia = "SELECT concat(vta.coddro,vta.nrocpb) as nrocpb, date_format(vta.feedoc,'%d/%m/%Y') as feedoc, date_format(vta.fevdoc,'%d/%m/%Y') as fevdoc, vta.codtdc as codtdc, vta.serdoc as serdoc, "
    sSentencia = sSentencia & "vta.nrodoc, vta.nrodoc_fin, aux.tpodci as tpodci, (CASE WHEN tpodci='01' THEN RIGHT(aux.rucaux, 8) ELSE aux.rucaux END) as codaux, trim(left(aux.razaux,60)) as razaux, "
    sSentencia = sSentencia & "ROUND(vta.impexp_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impexp, "
    sSentencia = sSentencia & "ROUND(vta.impogr_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impogr, "
    sSentencia = sSentencia & "(CASE WHEN vta.categoriadoc<>'" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impexo, "
    sSentencia = sSentencia & "(CASE WHEN vta.categoriadoc='" & CategoriaDocumento.RetencionOtro & "' THEN ROUND(vta.impexo_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) ELSE 0 END) as impina, "
    sSentencia = sSentencia & "ROUND(vta.impisc_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impisc, "
    sSentencia = sSentencia & "ROUND(vta.impigv_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impigv, "
    sSentencia = sSentencia & "ROUND(vta.impoim_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as impoim, "
    sSentencia = sSentencia & "ROUND(vta.imptot_" & sMoneda & " * (case tdc.sgntdc when 0 then -1 else 1 end), 2) as imptot, "
'ini 2015-09-16 t.cam solo ME
    'sSentencia = sSentencia & "vta.imptcb as imptcb, "
    sSentencia = sSentencia & "CASE tpomon WHEN '" & TPOMON_EXT & "' THEN vta.imptcb ELSE 0 END imptcb, "
'fin 2015-09-16 t.cam solo ME
    sSentencia = sSentencia & "date_format(feedoc_ref,'%d/%m/%Y') as d1, codtdc_ref as d2, serdoc_ref as d3, nrodoc_ref as d4, '" & sMoneda & "' as Moneda "
    'ini 2014-05-25
    sSentencia = sSentencia & ",impfob_mn "
    'fin 2014-05-25
    sSentencia = sSentencia & "FROM covtadoc vta "
    sSentencia = sSentencia & "INNER JOIN tgaux aux on vta.codaux=aux.codaux and aux.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "INNER JOIN tgtdc tdc on vta.codemp=tdc.codemp and vta.codtdc=tdc.codtdc and tdc.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "WHERE vta.codemp='" & gsCodEmp & "' and vta.pdoano='" & gsAnoAct & "' AND mespvs='" & xqmes & "' "
    sSentencia = sSentencia & "ORDER BY vta.codtdc,vta.serdoc,vta.nrodoc"
    
'ini 2014-05-30 adicion 5.3 plan ctas
   Case 32      'Plan de cuentas
    sArchivo = "050300" & "00" & "1"
    'gsNivCta
    Dim xxniv1 As String
    Dim xxnivf As String
    xxniv1 = Left(gsNivCta, 1)
    xxnivf = Right(gsNivCta, 1)
    sSentencia = "SELECT"
    sSentencia = sSentencia & "      a.codcta , a.detcta, b.codcta codcta2, b.detcta detcta2 "
    sSentencia = sSentencia & "FROM cocta a "
    sSentencia = sSentencia & "LEFT JOIN cocta b "
    sSentencia = sSentencia & "    ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND LEFT(a.codcta," & xxniv1 & ")=LEFT(b.codcta," & xxniv1 & ") "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "    AND a.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "    AND LENGTH(a.codcta)=" & xxnivf
    sSentencia = sSentencia & "    AND  LENGTH(b.codcta)=" & xxniv1 & " "
    sSentencia = sSentencia & "ORDER BY codcta"
'fin 2014-05-30 adicion 5.3 plan ctas

   Case Else: Exit Sub
  End Select
  sMoneda = Choose(cboTpoMon.ListIndex + 1, "1", "2")
  sArchivo = "LE" & gsRUCEmp & gsAnoAct & Left(cmbEjercicio.Text, 2) & "00" & sArchivo & "0" & sMoneda & "1" & ".txt"
  
  On Error GoTo CancelaDialogo
  cdlMain.DialogTitle = "Grabar Archivo Como"
  cdlMain.CancelError = True
  cdlMain.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
  cdlMain.FileName = sArchivo
  cdlMain.DefaultExt = ".txt"
  cdlMain.Filter = "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*"
  cdlMain.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  ChDir App.path
  If MsgBox("¿ Estás Seguro de Generar Registro Electrónico? ", vbQuestion + vbYesNo) = vbYes Then
    sArchivo = cdlMain.FileName
    sExpresion = cdlMain.FileTitle
    ppArchivoElectronico sArchivo, sExpresion, sSentencia
    MsgBox TEXT_8008, vbInformation
  End If
  ChDrive Left$(App.path, 1)
  ChDir App.path

End Sub

Private Sub ppArchivoElectronico_2015_11_18(ByVal sArchivo As String, ByVal sNombreArchivo As String, ByVal sSentencia As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String, s_Expresion As String
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nRegistroAux As Long, nRegistroDeta As Long, nTamano As Integer
  Dim sAuxiliar As String, s_OldMessage As String
  Dim nSumatoriaTotal As Double
  
  ' selecciono informacion de proceso
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
    nRegistros = .RecordCount
  End With
  
  ' Creo objeto de archivo
  If nRegistros > 0 Then
    s_Expresion = Left(sNombreArchivo, 30) & "1" & Mid(sNombreArchivo, 32)
    sArchivo = Replace(sArchivo, sNombreArchivo, s_Expresion)
  End If
  Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
  Set potxtFileExp = pofsoFileExp.CreateTextFile(sArchivo, True)
  s_Caracter = "|"
  
  ' detalle de archivo
  Dim xxMesPvs As String
  xxMesPvs = Left(cmbEjercicio.Text, 2)
  'ini 2014-07-31 numero correlativo
  '2015-04-07 error desbordamiento Dim xNroCorr As Integer
  Dim xNroCorr As Long
  xNroCorr = 1
  'fin 2014-07-31 numero correlativo
  If Not (porstMRp.BOF And porstMRp.EOF) Then
    nRegistro = 0
    While Not porstMRp.EOF
      psRegistro = ""
      Select Case pnOpcion
       Case 5       ' 3.2 libro caja
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!codbco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!NroCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!TpoMonCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        nImporte_mn = porstMRp!cDebe
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = porstMRp!cHaber
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
       '2015-11-12 PLE rpt fmt electro archivo  Case 6       ' 3.2 libro caja
       Case 6, 7      ' 3.2 libro caja,
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
             
'ini 2015-11-12/18 PLE rpt fmt electro archivo
'        s_Expresion = "NoExiste"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2015-11-12/18 PLE rpt fmt electro archivo

        
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = porstMRp!rucaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = porstMRp!razAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'ini 2015-11-12/13 PLE rpt fmt electro archivo
       Case 8   '3.4 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
             
        s_Expresion = "NoExiste"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!rucaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!razAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
         s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2015-11-12/13 PLE rpt fmt electro archivo
       Case 9, 10 '3.5 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS - TERCEROS O CUENTA 17
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
             
        s_Expresion = "NoExiste"
        psRegistro = psRegistro & s_Expresion & s_Caracter

        
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!rucaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!razAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
         s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
       Case 11   '3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
         ' 3: segun nueva estructura numero correlativo de asiento contable
         s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
         psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
        
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!rucaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!razAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!codtdc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!serdoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!nrodoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
                
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
         s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
       Case 23, 24        ' libro diario, simplificado
        ' 1: periodo
'ini 2015-04-01 convierte un solo mes 12y00
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        If Left(cmbEjercicio.Text, 2) = "12" Then
'        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        Else
'        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        End If
'fin 2015-04-01 convierte un solo mes 12y00
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
         
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
         ' 3: segun nueva estructura numero correlativo de asiento contable
         '2014-07-31 numero correlativo s_Expresion = porstMRp!coddro & Right(Format(porstMRp!NroCpb, "000000"), 5)
         s_Expresion = gfCeros(Str(xNroCorr), 8, 0, "0")
         '2015-04-06 convierte un solo mes 12y00  psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         If ExistFieldInRS(porstMRp, "mespvs") Then
            psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         Else
            psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         End If
         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
    
        ' 3: plan de cuentas - constante
'  2014-05-29 sale re`plaza por gnCodPlaCata        s_Expresion = "01"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: fecha emisión
        s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        ' 7: movimiento debe
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: movimiento haber
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
        ' 11: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
        psRegistro = psRegistro & s_Caracter
        ' 12: segun nueva estructura Número correlativo utilizado en el Registro de Compras
        psRegistro = psRegistro & s_Caracter
        ' 13: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
        psRegistro = psRegistro & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
                
        ' 9: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: codigo libro - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codlib) Then
          s_Expresion = porstMRp!codlib
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'******************************************************************
       Case 27        ' libro mayor
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
         psRegistro = psRegistro & s_Expresion & s_Caracter
         
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
         ' 3: segun nueva estructura numero correlativo de asiento contable
         '2014-07-31 numero correlativo s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", porstMRp!coddro & Right(Format(porstMRp!NroCpb, "000000"), 5))
         s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
         '2015-04-06 convierte un solo mes 12y00 psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
    
        ' 3: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: fecha emisión
        s_Expresion = "01/" & Left(cmbEjercicio.Text, 2) & "/" & gsAnoAct
        If Not IsNull(porstMRp!fehope) Then
          s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        ' 6: saldo o movimiento deudor
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 7: saldo o movimiento acreedor
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
        ' 11: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
        psRegistro = psRegistro & s_Caracter
        ' 12: segun nueva estructura Número correlativo utilizado en el Registro de Compras
        psRegistro = psRegistro & s_Caracter
        ' 13: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
        psRegistro = psRegistro & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
        
        ' 8: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 9: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: medio pago - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!medpago) Then
          s_Expresion = porstMRp!medpago
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
       Case 29, 30      ' registro compras
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actualiza libro electronico
        ' 3: Nueva version Contribuyentes del Régimen General: Número
        '2014-07-31 numero correlativo s_Expresion = "M" & Left(porstMRp!NroCpb, 4) & Right(porstMRp!NroCpb, 5)
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'ini 2015-05-23 actuliza libro electronico
        
        ' 4: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        s_Expresion = IIf(Format(porstMRp!codtdc, "00") = "14", s_Expresion, "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 7: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        s_Expresion = IIf(porstMRp!codtdc = "05", Right(porstMRp!serdoc, 1), porstMRp!serdoc) '2014-07-16
       
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: año emision DUA
        s_Expresion = IIf(IsNull(porstMRp!anno), "0", porstMRp!anno)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 9: numero comprobante de pago
        'ini 2015-05-26 actualiza libro electronico
        's_Expresion = porstMRp!nrodoc
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
        IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        '2014-08-20 Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), porstMRp!nrodoc)
        'fin 2015-05-26 actualiza libro electronico
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: numero final no dan derecho a credito - constante
        s_Expresion = "0"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: tipo documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: numero documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "-", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 14: base imponible adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp1), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 15: impuesto adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp2), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 16: base imponible adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp3), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 17: impuesto adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp4), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 18: base imponible adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp5), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 19: impuesto adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp6), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 20: adquisiciones no gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp7), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 21: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp8), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 22: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp9), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 23: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp10), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 24: tipo de cambio
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!ImpTCb), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 25: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!feedoc_ref) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!feedoc_ref, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 26: tipo comprobante pago modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 27: serie comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!serdoc_ref), "-", porstMRp!serdoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'      txtDetalle(0).MaxLength = .uorstMain!codaduana.DefinedSize
'      txtDetalle(1).MaxLength = .uorstMain!annodua.DefinedSize
'      txtDetalle(2).MaxLength = .uorstMain!nrodua.DefinedSize

        'ini 2015-05-23 actuliza libro electronico
        ' 28: Nueva version Contribuyentes del Régimen General: Número
        '2014-07-10 cambiar s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana) & IIf(IsNull(porstMRp!annodua), "", porstMRp!annodua) & IIf(IsNull(porstMRp!nrodua), "", porstMRp!nrodua)
        s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'ini 2015-05-23 actuliza libro electronico
        
        ' 28: numero comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!nrodoc_ref), "-", porstMRp!nrodoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 29: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!v1), "-", porstMRp!v1)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 30: fecha emision constancia detracción
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!FehCDt) Or s_Expresion = "0"), "01/01/0001", Format(porstMRp!FehCDt, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 31: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 32: comprobante afecto retencion
        s_Expresion = IIf(IsNull(porstMRp!indreten), "0", porstMRp!indreten)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 33: identifica ajuste - constante
        s_Expresion = Mid(Format(porstMRp!feedoc, "dd/mm/yyyy"), 4, 2)
        
        '2015-02-19 s_Expresion = IIf(s_Expresion = gsMesAct, "1", "6")
        's_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
        
        'ini 215-08-10 adicionado por teo
'        If (porstMRp!codtdc = "05" Or porstMRp!codtdc = "12") And (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "1", IIf(s_Expresion = gsMesAct, "1", "6"))
'        Else
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
'        End If

'        If (porstMRp!codtdc = "05" Or porstMRp!codtdc = "12") And (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "1", IIf(s_Expresion = gsMesAct, "1", "6"))
'        Else
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
'        End If
'        'ini 2015-08-10 correcion rafael fope-femision
'        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 12 Then
'            s_Expresion = "6"
'        End If
'        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 12 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
'            s_Expresion = "7"
'        End If
'ini 2015-08-26
          Select Case porstMRp!codtdc
                Case "05", "06", "07", "08", "11", "12", "13", "14"
                     If s_Expresion = gsMesAct Then
                        s_Expresion = "1"
                     Else
                        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                           s_Expresion = "6"
                        End If
                        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                           s_Expresion = "7"
                        End If
                     End If
                Case Else
                     If (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
                        s_Expresion = "0"
                     Else
                        If s_Expresion = gsMesAct Then
                           s_Expresion = "1"
                        Else
                           If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                              s_Expresion = "6"
                           End If
                           If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                              s_Expresion = "7"
                           End If
                        End If
                     End If
          End Select
'fin 2015-08-26
      

        
       'fin 2015-08-10 correcion rafael fope-femision
       'fin 215-08-10 adicionado por teo
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 34: Campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
       Case 31    ' resgistro ventas
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actuliza libro electronico
        ' 3: Version nueva
'''        'Contribuyentes del Régimen General: Número correlativo del asiento contable
'''        'identificado en el campo 2, cuando se utilice el Código Único de la Operación
'''        '(CUO). El primer dígito debe ser: "A" para el asiento de apertura del
'''        'ejercicio, "M" para los asientos de movimientos o ajustes del mes o "C" para
'''        'el asiento de cierre del ejercicio.
'''        '2. Contribuyentes del Régimen Especial de Renta - RER:  Número correlativo.
'''        'El primer dígito debe ser: "M".
        's_Expresion = "M" & porstMRp!NroCpb
        '2014-07-31 numero correlativo s_Expresion = "M" & Left(porstMRp!NroCpb, 4) & Right(porstMRp!NroCpb, 5)
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'fin 2015-05-23 actuliza libro electronico
              
        ' 3: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 7: numero comprobante de pago
        'ini 2015-05-26 actualiza libro electronico
        's_Expresion = porstMRp!nrodoc
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), porstMRp!nrodoc)
        'fin 2015-05-26 actualiza libro electronico
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: numero final agrupar documentos
        s_Expresion = IIf(IsNull(porstMRp!nrodoc_fin), "0", porstMRp!nrodoc_fin)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 9: tipo documento identidad cliente
        s_Expresion = Right(IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci), 1)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: numero documento identidad cliente
'ini 2014-07-10 validacion ruc 8dig tdoc=01
        's_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        If IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci) = "01" Then
             s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
             s_Expresion = Right(s_Expresion, 8)
       Else
            s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        End If
'ini 2014-07-10 validacion ruc 8dig tdoc=01
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "-", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: valor facturado exportacion
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impexp), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: base imponible operacion gravada
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impogr), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 14: importe total operacion exonerada
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impexo), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 15: importe total operacion inafecta
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impina), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 16: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impisc), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 17: igv y/o ipm
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impigv), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 18: base imponible operacion gravada ivap - constante
        s_Expresion = "0.00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 19: impuesto ventas arroz pilado (ivap) - constante
        s_Expresion = "0.00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 20: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impoim), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 21: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imptot), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 22: tipo de cambio
        s_Expresion = Replace(FormatNumber(CDec(IIf(IsNull(porstMRp!ImpTCb), 0, porstMRp!ImpTCb)), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 23: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!d2), "0", porstMRp!d2), "00")
        s_Expresion = IIf((IsNull(porstMRp!d1) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!d1, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 24: tipo comprobante pago modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!d2), "0", porstMRp!d2), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 25: serie comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!d3), "-", porstMRp!d3)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 26: numero comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!d4), "-", porstMRp!d4)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actuliza libro electronico
        ' 27: version nueva  Valor FOB embarcado de la exportación
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impfob_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'fin 2015-05-23 actuliza libro electronico
        
        ' 28: identifica estado comprobante periodo - constante
        s_Expresion = IIf(CDec(porstMRp!imptot) = 0, "2", "1")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 29: campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
        
'ini 2014-05-30 adicion 5.3 plan ctas
       Case 32        ' plan de cuentas
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & Format(Day(gfUltDia("01/" & Left(cmbEjercicio.Text, 2) & "/" & gsAnoAct)), "00") ' Format(Str(Day(Now)), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: Código de la Cuenta Contable desagregada hasta el nivel máximo de dígitos utilizado
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 3: Descripción de la Cuenta Contable desagregada al nivel máximo de dígitos utilizado
        s_Expresion = porstMRp!detcta
        psRegistro = psRegistro & s_Expresion & s_Caracter

        ' 4: Código del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = gnCodPlaCata
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: Descripción del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = IIf(gnCodPlaCata <> "99", "-", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter

'fin 2014-05-30 adicion 5.3 plan ctas
      End Select
      potxtFileExp.WriteLine psRegistro
      nRegistro = nRegistro + 1
      porstMRp.MoveNext
      xNroCorr = xNroCorr + 1 '2014-07-31 numero correlativo
    Wend
    ' Cierro objeto y saco de memoria
    potxtFileExp.Close
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  GoTo Finalizar
  
error:
Finalizar:
  ' Reinicializo los mensajes
  ' Coloco el puntero en normal

End Sub
Private Sub ppArchivoElectronico_2016_01_22(ByVal sArchivo As String, ByVal sNombreArchivo As String, ByVal sSentencia As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String, s_Expresion As String
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim nRegistroAux As Long, nRegistroDeta As Long, nTamano As Integer
  Dim sAuxiliar As String, s_OldMessage As String
  Dim nSumatoriaTotal As Double
  
   Dim n_SdoDeb As Double, n_SdoHab As Double 'teo 2015-12-15 falta definir
 
  ' selecciono informacion de proceso
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
    nRegistros = .RecordCount
  End With
  
    Select Case pnOpcion
    Case 19
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trpRngBceCpb", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb")
    End Select

  
  ' Creo objeto de archivo
  If nRegistros > 0 Then
    s_Expresion = Left(sNombreArchivo, 30) & "1" & Mid(sNombreArchivo, 32)
    sArchivo = Replace(sArchivo, sNombreArchivo, s_Expresion)
  End If
  Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
  Set potxtFileExp = pofsoFileExp.CreateTextFile(sArchivo, True)
  s_Caracter = "|"
  
  ' detalle de archivo
  Dim xxMesPvs As String
  xxMesPvs = Left(cmbEjercicio.Text, 2)
  'ini 2014-07-31 numero correlativo
  '2015-04-07 error desbordamiento Dim xNroCorr As Integer
  Dim xNroCorr As Long
  xNroCorr = 1
  'fin 2014-07-31 numero correlativo
  If Not (porstMRp.BOF And porstMRp.EOF) Then
    nRegistro = 0
    While Not porstMRp.EOF
      psRegistro = ""
      Select Case pnOpcion
       Case 5       ' 3.2 libro caja
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!codbco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!NroCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!TpoMonCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        nImporte_mn = cDebe
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = cHaber
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
      
       '2015-11-12 PLE rpt fmt electro archivo  Case 6       ' 3.2 libro caja
       '2015-11-12/18 PLE rpt fmt electro archivo  Case 6, 7      ' 3.2 libro caja,
'case=6y7 3.3 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 12 CUENTAS POR COBRAR COMERCIALES – TERCEROS Y 13
'case=8 3.4 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 14 CUENTAS POR COBRAR AL PERSONAL, A LOS ACCIONISTAS
'case=9y10 3.5 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 16 CUENTAS POR COBRAR DIVERSAS - TERCEROS O CUENTA 17
'case=11 3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
'case=12 3.11 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 41 REMUNERACIONES Y PARTICIPACIONES POR PAGAR (PCGE) (2)
'case=13y14 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=15y16 3.12 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 42 CUENTAS POR PAGAR COMERCIALES – TERCEROS Y LA CUENTA 43
'case=18 3.15  LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 37 ACTIVO DIFERIDO Y DE LA CUENTA 49 PASIVO DIFERIDO (PCGE)   (2)

'archivo agrupado y ordenado por RUC
       Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
             
'ini 2015-11-12/18 PLE rpt fmt electro archivo
'        s_Expresion = "NoExiste"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'fin 2015-11-12/18 PLE rpt fmt electro archivo
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 12
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
'fin 2015-12-07 libros inven y balance correcc
        s_Expresion = porstMRp!TpoDci
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = porstMRp!rucaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 13, 14, 15, 16
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
'fin 2015-12-07 libros inven y balance correcc
        
'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 12
        s_Expresion = porstMRp!codaux
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
        
'fin 2015-12-07 libros inven y balance correcc
        s_Expresion = porstMRp!razAux
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2015-12-07 libros inven y balance correcc
''        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
''        psRegistro = psRegistro & s_Expresion & s_Caracter
        Select Case pnOpcion
        Case 12
        Case Else
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
 'fin 2015-12-07 libros inven y balance correcc
       
 'ini 2015-12-07 libros inven y balance correcc
        Select Case pnOpcion
        Case 15, 16
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        End Select
'fin 2015-12-07 libros inven y balance correcc
      
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'ini 2015-12-04 libros inven y balance correcc
        s_Expresion = porstMRp!codtdc
       'psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & "-" & porstMRp!serdoc
       'psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & "-" & porstMRp!nrodoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
       Case 18
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
'ini 2015-11-12/18 PLE rpt fmt electro archivo
'        s_Expresion = "NoExiste"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
        psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
'fin 2015-11-12/18 PLE rpt fmt electro archivo
'ini 2015-12-04 libros inven y balance correcc
        s_Expresion = porstMRp!codtdc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & porstMRp!serdoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = s_Expresion & porstMRp!nrodoc
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        s_Expresion = porstMRp!GloIte
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        nImporte_mn = CDec(porstMRp!DebeSol)
        nImporte_me = CDec(porstMRp!HaberSol)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'adiciones
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        'deducciones
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
        
'ini 2015-11-12/18 PLE rpt fmt electro archivo


'fin 2015-12-04 libros inven y balance correcc
        
'ini 2015-11-12/13 PLE rpt fmt electro archivo
        
'''''ini 2015-12-07 libros inven y balance correcc
''''       Case 11   '3.6 LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 ESTIMACIÓN DE CUENTAS DE COBRANZA DUDOSA (PCGE) (2)
''''        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        ' 2: numero correlativo o codigo unico
''''        s_Expresion = "0000-000000-000000"
''''        If Not IsNull(porstMRp!coddro) Then
''''          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
''''        End If
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
''''         ' 3: segun nueva estructura numero correlativo de asiento contable
''''         s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
''''         psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
''''    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
''''
''''        s_Expresion = porstMRp!TpoDci
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!rucaux
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!razAux
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!codtdc
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!serdoc
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = porstMRp!nrodoc
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''        nImporte_mn = CDec(porstMRp!DebeSol)
''''        nImporte_me = CDec(porstMRp!HaberSol)
''''        n_Importe = Round(nImporte_mn - nImporte_me, 2)
''''        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
''''
''''         s_Expresion = "1"
''''        psRegistro = psRegistro & s_Expresion & s_Caracter
'''''fin 2015-12-07 libros inven y balance correcc
        
'ini 2015-12-10 PLE rpt fmt electro archivo
       Case 19
       'Periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Código de la cuenta
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Saldos iniciales Debe
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cApeD), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Saldos iniciales Haber
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cApeH), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Movimientos del ejercicio o periodo - Debe
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cSumaD), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Movimientos del ejercicio o periodo - Haber
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!cSumaH), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'teo 2015-12-15 falta definir
        psRegistro = psRegistro & "falta suma del mayo debe" & s_Caracter
        psRegistro = psRegistro & "falta suma del mayo haber" & s_Caracter
        'teo 2015-12-15 falta definir
        
       'Saldos al 31 de Diciembre - Deudor (saldos finales deudor)
        nImporte_mn = CDec(porstMRp!cApeD + cSumaD)
        nImporte_me = CDec(porstMRp!cApeH + cSumaH)
        n_Importe = Round(IIf(nImporte_mn > nImporte_me, nImporte_mn - nImporte_me, 0), 2)
        n_SdoDeb = n_Importe 'teo 2015-12-15 falta definir
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'Saldos al 31 de Diciembre - Acreedor (saldos finales deudor)
        nImporte_mn = CDec(porstMRp!cApeD + cSumaD)
        nImporte_me = CDec(porstMRp!cApeH + cSumaH)
        n_Importe = Round(IIf(nImporte_mn > nImporte_me, 0, nImporte_mn - nImporte_me), 2)
        n_SdoHab = n_Importe 'teo 2015-12-15 falta definir
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'teo 2015-12-15 falta definir
        psRegistro = psRegistro & "falta Transferencias y Cancelaciones - Debe" & s_Caracter
        psRegistro = psRegistro & "falta Transferencias y Cancelaciones - Haber" & s_Caracter
        psRegistro = psRegistro & "falta Cuentas de Balance - Activo" & s_Caracter
        psRegistro = psRegistro & "falta Cuentas de Balance - Pasivo" & s_Caracter
        'teo 2015-12-15 falta definir
        
        'Resultado por Naturaleza - Pérdidas / Sdo. Finales del estado de perdidas y ganancia por funcion peridida
        n_Importe = Round(IIf(n_SdoDeb > 0 And (porstMRp!TpoSdo = "F" Or porstMRp!TpoSdo = "A"), n_SdoDeb, 0), 2)
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'Resultado por Naturaleza - Ganancias / Sdo. Finales del estado de perdidas y ganancia por funcion ganancia
        n_Importe = Round(IIf(n_SdoHab > 0 And (porstMRp!TpoSdo = "F" Or porstMRp!TpoSdo = "A"), n_SdoHab, 0), 2)
        s_Expresion = Replace(FormatNumber(CDec(n_Importe), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
       
        'teo 2015-12-15 falta definir
        psRegistro = psRegistro & "falta Adiciones" & s_Caracter
        psRegistro = psRegistro & "falta Deducciones" & s_Caracter
        'teo 2015-12-15 falta definir
       
        '2015-11-12/18 PLE rpt fmt electro archivo s_Expresion = "1"
        s_Expresion = IIf(Left(cmbEjercicio.Text, 2) = "00", "9", "1")
       
'fin 2015-12-10 PLE rpt fmt electro archivo
       Case 23, 24        ' libro diario, simplificado
        ' 1: periodo
'ini 2015-04-01 convierte un solo mes 12y00
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        If Left(cmbEjercicio.Text, 2) = "12" Then
'        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        Else
'        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
'        End If
'fin 2015-04-01 convierte un solo mes 12y00
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
         
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
         ' 3: segun nueva estructura numero correlativo de asiento contable
         '2014-07-31 numero correlativo s_Expresion = porstMRp!coddro & Right(Format(porstMRp!NroCpb, "000000"), 5)
         s_Expresion = gfCeros(Str(xNroCorr), 8, 0, "0")
         '2015-04-06 convierte un solo mes 12y00  psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         If ExistFieldInRS(porstMRp, "mespvs") Then
            psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         Else
            psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         End If
         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
    
        ' 3: plan de cuentas - constante
'  2014-05-29 sale re`plaza por gnCodPlaCata        s_Expresion = "01"
'        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: fecha emisión
        s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        ' 7: movimiento debe
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: movimiento haber
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
        ' 11: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
        psRegistro = psRegistro & s_Caracter
        ' 12: segun nueva estructura Número correlativo utilizado en el Registro de Compras
        psRegistro = psRegistro & s_Caracter
        ' 13: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
        psRegistro = psRegistro & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
                
        ' 9: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: codigo libro - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codlib) Then
          s_Expresion = porstMRp!codlib
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'******************************************************************
       Case 27        ' libro mayor
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = "0000-000000-000000"
        If Not IsNull(porstMRp!coddro) Then
          s_Expresion = porstMRp!coddro & "-" & Format(porstMRp!NroCpb, "000000") & "-" & Format(porstMRp!NroIte, "000000")
        End If
         psRegistro = psRegistro & s_Expresion & s_Caracter
         
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
         ' 3: segun nueva estructura numero correlativo de asiento contable
         '2014-07-31 numero correlativo s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", porstMRp!coddro & Right(Format(porstMRp!NroCpb, "000000"), 5))
         s_Expresion = IIf(IsNull(porstMRp!coddro), "000000000", gfCeros(Str(xNroCorr), 8, 0, "0"))
         '2015-04-06 convierte un solo mes 12y00 psRegistro = psRegistro & IIf(xxMesPvs = "00", "A", IIf(xxMesPvs = "13", "C", "M")) & s_Expresion & s_Caracter
         psRegistro = psRegistro & IIf(porstMRp!mespvs = "00", "A", IIf(porstMRp!mespvs = "13", "C", "M")) & s_Expresion & s_Caracter
         ' 4: segun nueva estructura Código del Plan de Cuentas utilizado por el deudor tributario
         psRegistro = psRegistro & gnCodPlaCata & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
    
        ' 3: cuenta contable
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: fecha emisión
        s_Expresion = "01/" & Left(cmbEjercicio.Text, 2) & "/" & gsAnoAct
        If Not IsNull(porstMRp!fehope) Then
          s_Expresion = Format(porstMRp!fehope, "dd/mm/yyyy")
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: glosa o descripción
        s_Expresion = "-"
        If Not IsNull(porstMRp!GloIte) Then
          s_Expresion = gfSacaEntRetApos(porstMRp!GloIte)
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' importes
        nImporte_mn = CDec(porstMRp!nDebe)
        nImporte_me = CDec(porstMRp!nHaber)
        n_Importe = Round(nImporte_mn - nImporte_me, 2)
        ' 6: saldo o movimiento deudor
        nImporte_mn = Abs(IIf(n_Importe >= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 7: saldo o movimiento acreedor
        nImporte_me = Abs(IIf(n_Importe <= 0, n_Importe, 0))
        s_Expresion = Replace(FormatNumber(CDec(nImporte_me), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
    'ini 2014-05-29 adicion nuevas coluimna libros electronicos
        ' 11: segun nueva estructura Número correlativo utilizado en el Registro de Ventas e Ingresos.
        psRegistro = psRegistro & s_Caracter
        ' 12: segun nueva estructura Número correlativo utilizado en el Registro de Compras
        psRegistro = psRegistro & s_Caracter
        ' 13: segun nueva estructura Número correlativo utilizado en el Registro de Consignaciones
        psRegistro = psRegistro & s_Caracter
    'fin 2014-05-29 adicion nuevas coluimna libros electronicos
        
        ' 8: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 9: documento - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!cDocume) Then
          s_Expresion = porstMRp!cDocume
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: auxiliar - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!codaux) Then
          s_Expresion = porstMRp!codaux & "-" & porstMRp!razAux
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: referencia - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!RefDoc) Then
          s_Expresion = porstMRp!RefDoc
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: medio pago - opcional
        s_Expresion = "-"
        If Not IsNull(porstMRp!medpago) Then
          s_Expresion = porstMRp!medpago
        End If
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
       Case 29, 30      ' registro compras
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actualiza libro electronico
        ' 3: Nueva version Contribuyentes del Régimen General: Número
        '2014-07-31 numero correlativo s_Expresion = "M" & Left(porstMRp!NroCpb, 4) & Right(porstMRp!NroCpb, 5)
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'ini 2015-05-23 actuliza libro electronico
        
        ' 4: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        s_Expresion = IIf(Format(porstMRp!codtdc, "00") = "14", s_Expresion, "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 7: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        s_Expresion = IIf(porstMRp!codtdc = "05", Right(porstMRp!serdoc, 1), porstMRp!serdoc) '2014-07-16
       
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: año emision DUA
        s_Expresion = IIf(IsNull(porstMRp!anno), "0", porstMRp!anno)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 9: numero comprobante de pago
        'ini 2015-05-26 actualiza libro electronico
        's_Expresion = porstMRp!nrodoc
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), _
        IIf(porstMRp!codtdc = "36", Right(porstMRp!nrodoc, 8), porstMRp!nrodoc))
        '2014-08-20 Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), porstMRp!nrodoc)
        'fin 2015-05-26 actualiza libro electronico
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: numero final no dan derecho a credito - constante
        s_Expresion = "0"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: tipo documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: numero documento identidad proveedor
        s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "-", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 14: base imponible adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp1), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 15: impuesto adquisiciones gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp2), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 16: base imponible adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp3), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 17: impuesto adquisiciones gravadas dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp4), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 18: base imponible adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp5), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 19: impuesto adquisiciones gravadas no dan derecho
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp6), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 20: adquisiciones no gravadas
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp7), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 21: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp8), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 22: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp9), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 23: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imp10), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 24: tipo de cambio
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!ImpTCb), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 25: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!feedoc_ref) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!feedoc_ref, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 26: tipo comprobante pago modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00")
        s_Expresion = IIf(s_Expresion = "91", "00", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 27: serie comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!serdoc_ref), "-", porstMRp!serdoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
'      txtDetalle(0).MaxLength = .uorstMain!codaduana.DefinedSize
'      txtDetalle(1).MaxLength = .uorstMain!annodua.DefinedSize
'      txtDetalle(2).MaxLength = .uorstMain!nrodua.DefinedSize

        'ini 2015-05-23 actuliza libro electronico
        ' 28: Nueva version Contribuyentes del Régimen General: Número
        '2014-07-10 cambiar s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana) & IIf(IsNull(porstMRp!annodua), "", porstMRp!annodua) & IIf(IsNull(porstMRp!nrodua), "", porstMRp!nrodua)
        s_Expresion = IIf(IsNull(porstMRp!codaduana), "", porstMRp!codaduana)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'ini 2015-05-23 actuliza libro electronico
        
        ' 28: numero comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!nrodoc_ref), "-", porstMRp!nrodoc_ref)
        s_Expresion = IIf(Format(IIf(IsNull(porstMRp!codtdc_ref), "0", porstMRp!codtdc_ref), "00") = "91", "-", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 29: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!v1), "-", porstMRp!v1)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 30: fecha emision constancia detracción
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        s_Expresion = IIf((IsNull(porstMRp!FehCDt) Or s_Expresion = "0"), "01/01/0001", Format(porstMRp!FehCDt, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 31: numero comprobante pago no domiciliado
        s_Expresion = IIf(IsNull(porstMRp!NroCDt), "0", porstMRp!NroCDt)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 32: comprobante afecto retencion
        s_Expresion = IIf(IsNull(porstMRp!indreten), "0", porstMRp!indreten)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 33: identifica ajuste - constante
        s_Expresion = Mid(Format(porstMRp!feedoc, "dd/mm/yyyy"), 4, 2)
        
        '2015-02-19 s_Expresion = IIf(s_Expresion = gsMesAct, "1", "6")
        's_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
        
        'ini 215-08-10 adicionado por teo
'        If (porstMRp!codtdc = "05" Or porstMRp!codtdc = "12") And (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "1", IIf(s_Expresion = gsMesAct, "1", "6"))
'        Else
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
'        End If

'        If (porstMRp!codtdc = "05" Or porstMRp!codtdc = "12") And (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "1", IIf(s_Expresion = gsMesAct, "1", "6"))
'        Else
'            s_Expresion = IIf((porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0#, "0", IIf(s_Expresion = gsMesAct, "1", "6"))
'        End If
'        'ini 2015-08-10 correcion rafael fope-femision
'        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 12 Then
'            s_Expresion = "6"
'        End If
'        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 12 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
'            s_Expresion = "7"
'        End If
'ini 2015-08-26
          Select Case porstMRp!codtdc
                Case "05", "06", "07", "08", "11", "12", "13", "14"
                     If s_Expresion = gsMesAct Then
                        s_Expresion = "1"
                     Else
                        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                           s_Expresion = "6"
                        End If
                        If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                           s_Expresion = "7"
                        End If
                     End If
                Case Else
                     If (porstMRp!imp2 + porstMRp!imp4 + porstMRp!imp6) = 0# Then
                        s_Expresion = "0"
                     Else
                        If s_Expresion = gsMesAct Then
                           s_Expresion = "1"
                        Else
                           If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) <= 12 Then
                              s_Expresion = "6"
                           End If
                           If Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) >= 13 Or Abs(DateDiff("m", porstMRp!fehope, porstMRp!feedoc)) = 13 Then
                              s_Expresion = "7"
                           End If
                        End If
                     End If
          End Select
'fin 2015-08-26
      

        
       'fin 2015-08-10 correcion rafael fope-femision
       'fin 215-08-10 adicionado por teo
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 34: Campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
       Case 31    ' resgistro ventas
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & "00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: numero correlativo o codigo unico
        s_Expresion = porstMRp!NroCpb
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actuliza libro electronico
        ' 3: Version nueva
'''        'Contribuyentes del Régimen General: Número correlativo del asiento contable
'''        'identificado en el campo 2, cuando se utilice el Código Único de la Operación
'''        '(CUO). El primer dígito debe ser: "A" para el asiento de apertura del
'''        'ejercicio, "M" para los asientos de movimientos o ajustes del mes o "C" para
'''        'el asiento de cierre del ejercicio.
'''        '2. Contribuyentes del Régimen Especial de Renta - RER:  Número correlativo.
'''        'El primer dígito debe ser: "M".
        's_Expresion = "M" & porstMRp!NroCpb
        '2014-07-31 numero correlativo s_Expresion = "M" & Left(porstMRp!NroCpb, 4) & Right(porstMRp!NroCpb, 5)
        s_Expresion = "M" & gfCeros(Str(xNroCorr), 8, 0, "0")
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'fin 2015-05-23 actuliza libro electronico
              
        ' 3: fecha emisión
        s_Expresion = Format(porstMRp!feedoc, "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 4: fecha vencimiento o pago
        s_Expresion = Format(IIf(IsNull(porstMRp!fevdoc), "", porstMRp!fevdoc), "dd/mm/yyyy")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: tipo comprobante de pago
        s_Expresion = Format(porstMRp!codtdc, "00")
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: serie comprobante de pago
        s_Expresion = IIf(IsNull(porstMRp!serdoc), "-", porstMRp!serdoc)
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 7: numero comprobante de pago
        'ini 2015-05-26 actualiza libro electronico
        's_Expresion = porstMRp!nrodoc
        s_Expresion = IIf(porstMRp!codtdc = "01" Or porstMRp!codtdc = "03" _
        Or porstMRp!codtdc = "04" Or porstMRp!codtdc = "06" _
        Or porstMRp!codtdc = "07" Or porstMRp!codtdc = "08", Right(porstMRp!nrodoc, 7), porstMRp!nrodoc)
        'fin 2015-05-26 actualiza libro electronico
        's_Expresion = IIf(porstMRp!imptot = 0, "", s_Expresion) '2014-07-10 validacion tot.fac=0 blanco
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 8: numero final agrupar documentos
        s_Expresion = IIf(IsNull(porstMRp!nrodoc_fin), "0", porstMRp!nrodoc_fin)
        s_Expresion = IIf(s_Expresion = "", "0", s_Expresion)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 9: tipo documento identidad cliente
        s_Expresion = Right(IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci), 1)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 10: numero documento identidad cliente
'ini 2014-07-10 validacion ruc 8dig tdoc=01
        's_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        If IIf(IsNull(porstMRp!TpoDci), "0", porstMRp!TpoDci) = "01" Then
             s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
             s_Expresion = Right(s_Expresion, 8)
       Else
            s_Expresion = IIf(IsNull(porstMRp!codaux), "-", porstMRp!codaux)
        End If
'ini 2014-07-10 validacion ruc 8dig tdoc=01
        
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 11: razon social proveedor
        s_Expresion = IIf(IsNull(porstMRp!razAux), "-", porstMRp!razAux)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 12: valor facturado exportacion
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impexp), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 13: base imponible operacion gravada
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impogr), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 14: importe total operacion exonerada
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impexo), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 15: importe total operacion inafecta
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impina), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 16: impuesto selectivo al consumo
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impisc), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 17: igv y/o ipm
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impigv), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 18: base imponible operacion gravada ivap - constante
        s_Expresion = "0.00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 19: impuesto ventas arroz pilado (ivap) - constante
        s_Expresion = "0.00"
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 20: otros tributos
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impoim), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 21: importe total
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!imptot), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 22: tipo de cambio
        s_Expresion = Replace(FormatNumber(CDec(IIf(IsNull(porstMRp!ImpTCb), 0, porstMRp!ImpTCb)), 3), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 23: fecha emision comprobante modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!d2), "0", porstMRp!d2), "00")
        s_Expresion = IIf((IsNull(porstMRp!d1) Or s_Expresion = "00"), "01/01/0001", Format(porstMRp!d1, "dd/mm/yyyy"))
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 24: tipo comprobante pago modifica
        s_Expresion = Format(IIf(IsNull(porstMRp!d2), "0", porstMRp!d2), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 25: serie comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!d3), "-", porstMRp!d3)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 26: numero comprobante pago modifica
        s_Expresion = IIf(IsNull(porstMRp!d4), "-", porstMRp!d4)
        psRegistro = psRegistro & s_Expresion & s_Caracter
        
        'ini 2015-05-23 actuliza libro electronico
        ' 27: version nueva  Valor FOB embarcado de la exportación
        s_Expresion = Replace(FormatNumber(CDec(porstMRp!impfob_mn), 2), ",", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        'fin 2015-05-23 actuliza libro electronico
        
        ' 28: identifica estado comprobante periodo - constante
        s_Expresion = IIf(CDec(porstMRp!imptot) = 0, "2", "1")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 29: campos libre - constante
        s_Expresion = ""
        psRegistro = psRegistro & s_Expresion & s_Caracter
'*******************************************************************************
        
'ini 2014-05-30 adicion 5.3 plan ctas
       Case 32        ' plan de cuentas
        ' 1: periodo
        s_Expresion = gsAnoAct & Left(cmbEjercicio.Text, 2) & Format(Day(gfUltDia("01/" & Left(cmbEjercicio.Text, 2) & "/" & gsAnoAct)), "00") ' Format(Str(Day(Now)), "00")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 2: Código de la Cuenta Contable desagregada hasta el nivel máximo de dígitos utilizado
        s_Expresion = porstMRp!CodCta
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 3: Descripción de la Cuenta Contable desagregada al nivel máximo de dígitos utilizado
        s_Expresion = porstMRp!detcta
        psRegistro = psRegistro & s_Expresion & s_Caracter

        ' 4: Código del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = gnCodPlaCata
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 5: Descripción del Plan de Cuentas utilizado por el deudor tributario
        s_Expresion = IIf(gnCodPlaCata <> "99", "-", "")
        psRegistro = psRegistro & s_Expresion & s_Caracter
        ' 6: estado operación - contstante
        s_Expresion = "1"
        psRegistro = psRegistro & s_Expresion & s_Caracter

'fin 2014-05-30 adicion 5.3 plan ctas
      End Select
      potxtFileExp.WriteLine psRegistro
      nRegistro = nRegistro + 1
      porstMRp.MoveNext
      xNroCorr = xNroCorr + 1 '2014-07-31 numero correlativo
    Wend
    ' Cierro objeto y saco de memoria
    potxtFileExp.Close
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  GoTo Finalizar
  
error:
Finalizar:
  ' Reinicializo los mensajes
  ' Coloco el puntero en normal

End Sub


