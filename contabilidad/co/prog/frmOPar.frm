VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOPar 
   Caption         =   "[Entidad]"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1065
      ScaleHeight     =   690
      ScaleWidth      =   2955
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4260
      Width           =   2955
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
         Height          =   560
         Left            =   2220
         Picture         =   "frmOPar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   700
      End
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1485
         Picture         =   "frmOPar.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   700
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   780
         Picture         =   "frmOPar.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   700
      End
      Begin VB.CommandButton cmdCorregir 
         Caption         =   "&Corregir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   60
         Picture         =   "frmOPar.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   60
         Width           =   720
      End
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nivel de Cuentas"
      TabPicture(0)   =   "frmOPar.frx":0498
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkNiveles(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkNiveles(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkNiveles(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkNiveles(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkNiveles(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkNiveles(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkNivel2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Monedas"
      TabPicture(1)   =   "frmOPar.frx":04B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "cboMonFnc"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "txtDato(0)"
      Tab(1).Control(6)=   "txtDato(1)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Impuestos"
      TabPicture(2)   =   "frmOPar.frx":04D0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4(0)"
      Tab(2).Control(1)=   "Label5(0)"
      Tab(2).Control(2)=   "Label6(0)"
      Tab(2).Control(3)=   "Label7(0)"
      Tab(2).Control(4)=   "txtDato(2)"
      Tab(2).Control(5)=   "txtDato(3)"
      Tab(2).Control(6)=   "txtDato(4)"
      Tab(2).Control(7)=   "txtDato(5)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Retención/Percepción"
      TabPicture(3)   =   "frmOPar.frx":04EC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdDatoAyud(10)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdDatoAyud(9)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdDatoAyud(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdDatoAyud(6)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtDato(10)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtDato(11)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtDato(6)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtDato(7)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtDato(8)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtDato(9)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label6(2)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label7(2)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label4(1)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label5(1)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label6(1)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label7(1)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "Varios"
      TabPicture(4)   =   "frmOPar.frx":0508
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label4(2)"
      Tab(4).Control(1)=   "txtDato(12)"
      Tab(4).ControlCount=   2
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   10
         Left            =   -70830
         Picture         =   "frmOPar.frx":0524
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2830
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   9
         Left            =   -71490
         Picture         =   "frmOPar.frx":06CE
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2470
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   7
         Left            =   -70830
         Picture         =   "frmOPar.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1740
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   6
         Left            =   -71490
         Picture         =   "frmOPar.frx":0A22
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1400
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   12
         Left            =   -73320
         TabIndex        =   42
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox txtDato 
         Height          =   315
         Index           =   10
         Left            =   -71820
         TabIndex        =   39
         Top             =   2820
         Width           =   975
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   -71820
         TabIndex        =   38
         Top             =   3180
         Width           =   615
      End
      Begin VB.TextBox txtDato 
         Height          =   315
         Index           =   6
         Left            =   -71820
         TabIndex        =   33
         Top             =   1380
         Width           =   315
      End
      Begin VB.TextBox txtDato 
         Height          =   315
         Index           =   7
         Left            =   -71820
         TabIndex        =   32
         Top             =   1725
         Width           =   975
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   -71820
         TabIndex        =   31
         Top             =   2085
         Width           =   615
      End
      Begin VB.TextBox txtDato 
         Height          =   315
         HideSelection   =   0   'False
         Index           =   9
         Left            =   -71820
         TabIndex        =   30
         Top             =   2445
         Width           =   315
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   5
         Left            =   -71265
         TabIndex        =   20
         Top             =   2985
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   4
         Left            =   -71265
         TabIndex        =   19
         Top             =   2505
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   3
         Left            =   -71265
         TabIndex        =   18
         Top             =   2025
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   2
         Left            =   -71265
         TabIndex        =   17
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Height          =   375
         Index           =   1
         Left            =   -72600
         TabIndex        =   16
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox txtDato 
         Height          =   375
         Index           =   0
         Left            =   -72600
         TabIndex        =   15
         Top             =   3000
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Monedas de Trabajo"
         ForeColor       =   &H80000002&
         Height          =   795
         Left            =   -74520
         TabIndex        =   22
         Top             =   1920
         Width           =   3855
         Begin VB.OptionButton optMon2 
            Caption         =   "&2 Monedas"
            ForeColor       =   &H80000001&
            Height          =   195
            Left            =   2220
            TabIndex        =   14
            Top             =   360
            Width           =   1155
         End
         Begin VB.OptionButton optMon1 
            Caption         =   "&1 Moneda"
            ForeColor       =   &H80000001&
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.ComboBox cboMonFnc 
         Height          =   315
         Left            =   -72720
         TabIndex        =   12
         Top             =   1380
         Width           =   1335
      End
      Begin VB.CheckBox chkNivel2 
         Caption         =   "2 Dígitos"
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkNiveles 
         Caption         =   "8 Dígitos"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   2040
         TabIndex        =   11
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox chkNiveles 
         Caption         =   "7 Dígitos"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   10
         Top             =   3060
         Width           =   1215
      End
      Begin VB.CheckBox chkNiveles 
         Caption         =   "6 Dígitos"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkNiveles 
         Caption         =   "5 Dígitos"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   8
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CheckBox chkNiveles 
         Caption         =   "4 Dígitos"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkNiveles 
         Caption         =   "3 Dígitos"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe UIT"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   -74460
         TabIndex        =   43
         Top             =   1485
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   41
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   40
         Top             =   3240
         Width           =   1620
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   37
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   36
         Top             =   1785
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   35
         Top             =   2145
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   34
         Top             =   2505
         Width           =   2040
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto Extraordinario de Solidaridad"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   -74325
         TabIndex        =   29
         Top             =   3045
         Width           =   2700
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto a la Renta de 4ª Categoría"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   -74325
         TabIndex        =   28
         Top             =   2565
         Width           =   2595
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto Selectivo al Consumo"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   -74325
         TabIndex        =   27
         Top             =   2085
         Width           =   2220
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Impuesto General a las Ventas"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   -74325
         TabIndex        =   26
         Top             =   1605
         Width           =   2160
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda Extranjera"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -74100
         TabIndex        =   25
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda Nacional"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -74100
         TabIndex        =   24
         Top             =   3060
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda Funcional de la Empresa"
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   -74340
         TabIndex        =   23
         Top             =   1320
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmOPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pocnnMain As ADODB.Connection
Private porstCOCfg As ADODB.Recordset
Private porstTGCfg As ADODB.Recordset
Private pbNuevo As Boolean

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 6, 7, 9, 10
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub Form_Load()
   Dim dnContador As Integer

   Me.KeyPreview = True
   
 '[Recordsets                          'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstCOCfg = New ADODB.Recordset
   Set porstTGCfg = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstCOCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCta_Nv3, CodCta_Nv4, CodCta_Nv5, CodCta_Nv6, CodCta_Nv7, CodCta_Nv8," _
              & "  TpoMon_Fnc, TpoMon_Sgn_MN, TpoMon_Sgn_ME, IndMNE, CodTDc_Pcp, CodTDc_Rtc, CodCta_Pcp, CodCta_Rtc " _
              & "FROM COCfg"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With porstTGCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT PctIGV, PctISC, PctIR4, PctIES, PctRtc, PctPcp, ImpUIT " _
              & "FROM TGCfg"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
 ']

 '[Datos                               'Cambiar.
   With cboMonFnc
      .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
   End With

'   With chkNiveles
'      .Item(0).DataField = "CodCta_Nv3"
'      .Item(1).DataField = "CodCta_Nv4"
'      .Item(2).DataField = "CodCta_Nv5"
'      .Item(3).DataField = "CodCta_Nv6"
'      .Item(4).DataField = "CodCta_Nv7"
'      .Item(5).DataField = "CodCta_Nv8"
'      For dnContador = 0 To .Count - 1
'         Set .Item(dnContador).DataSource = porstMain
'      Next
'   End With
 ']
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   ppDatosDesconectados 1
   ppHabilitacion False
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'   Call gpTeclasData2(KeyAscii)
'End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   porstCOCfg.Close
   pocnnMain.Close
   Set porstCOCfg = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdCorregir_Click()
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   ppHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   If sstMain.Tab = 2 Then
      cboMonFnc.SetFocus
   Else
      chkNiveles(0).SetFocus
   End If
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   pocnnMain.BeginTrans                'INICIA TRANSACCION.
   
   ppDatosDesconectados 0
   
   porstCOCfg.Update
   porstTGCfg.Update
   
   pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
 
 '[Propio del formulario.
   With porstCOCfg
      gsNivCta = "2" & IIf(!CodCta_Nv3, "3", "") & IIf(!CodCta_Nv4, "4", "") & IIf(!CodCta_Nv5, "5", "") & IIf(!CodCta_Nv6, "6", "") & IIf(!CodCta_Nv7, "7", "") & IIf(!CodCta_Nv8, "8", "")
      gnIndMNE = IIf(IsNull(!IndMNE), 0, !IndMNE)
      gsTpoMon_Fnc = IIf(IsNull(!TpoMon_Fnc), "", !TpoMon_Fnc)
      gsTpoMon_Sgn_MN = IIf(IsNull(!TpoMon_Sgn_MN), "", !TpoMon_Sgn_MN)
      gsTpoMon_Sgn_ME = IIf(IsNull(!TpoMon_Sgn_ME), "", !TpoMon_Sgn_ME)
      gsCodTDc_Pcp = IIf(IsNull(!COdTDC_Pcp), "", !COdTDC_Pcp)
      gsCodTDc_Rtc = IIf(IsNull(!CodTDc_Rtc), "", !CodTDc_Rtc)
      gsCodCta_Pcp = IIf(IsNull(!COdCta_Pcp), "", !COdCta_Pcp)
      gsCodCta_Pcp = IIf(IsNull(!CodCta_Rtc), "", !CodCta_Rtc)
'gsCodTDc_Pcp As String, _
'gsCodTDc_Rtc As String
   End With
   With porstTGCfg
      gnPctIGV = CDec(!PctIGV)
      gnPctISC = CDec(!PctISC)
      gnPctIR4 = CDec(!PctIR4)
      gnPctIES = CDec(!PctIES)
      gnPctRtc = CDec(!PctRtc)
      gnPctPcp = CDec(!PctPcp)
      gnImpUIT = CDec(!ImpUIT)
   End With
 ']
   
   cmdCorregir.Enabled = True
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   ppHabilitacion False
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub cmdDeshacer_Click()
   On Error GoTo Err

   ppDatosDesconectados 1
   cmdCorregir.Enabled = True
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   ppHabilitacion False

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
   Case 6, 7, 9, 10
      If KeyCode = vbKeyF2 Then
         ppAyuBus Index
      End If
End Select
End Sub

Private Sub ppDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   If tnFase = 0 Then
     'Datos.
      porstCOCfg!TpoMon_Fnc = IIf(cboMonFnc.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      porstCOCfg!CodCta_Nv3 = IIf(chkNiveles(0).Value = vbChecked, 1, 0)
      porstCOCfg!CodCta_Nv4 = IIf(chkNiveles(1).Value = vbChecked, 1, 0)
      porstCOCfg!CodCta_Nv5 = IIf(chkNiveles(2).Value = vbChecked, 1, 0)
      porstCOCfg!CodCta_Nv6 = IIf(chkNiveles(3).Value = vbChecked, 1, 0)
      porstCOCfg!CodCta_Nv7 = IIf(chkNiveles(4).Value = vbChecked, 1, 0)
      porstCOCfg!CodCta_Nv8 = IIf(chkNiveles(5).Value = vbChecked, 1, 0)
'      porstMain!CodSoc = IIf(dcoSocio.BoundText = "", " ", dcoSocio.BoundText)
'      porstMain!FehOpe = dtpFecha.Value
      porstCOCfg!IndMNE = IIf(optMon1.Value, INDMNE_INA, INDMNE_ACT)
      porstCOCfg!TpoMon_Sgn_MN = txtDato(0).Text
      porstCOCfg!TpoMon_Sgn_ME = txtDato(1).Text
      porstCOCfg!COdTDC_Pcp = txtDato(9).Text
      porstCOCfg!CodTDc_Rtc = txtDato(6).Text
      porstCOCfg!COdCta_Pcp = txtDato(10).Text
      porstCOCfg!CodCta_Rtc = txtDato(7).Text
      porstTGCfg!PctIGV = CDec(txtDato(2).Text)
      porstTGCfg!PctISC = CDec(txtDato(3).Text)
      porstTGCfg!PctIR4 = CDec(txtDato(4).Text)
      porstTGCfg!PctIES = CDec(txtDato(5).Text)
      porstTGCfg!PctRtc = CDec(txtDato(8).Text)
      porstTGCfg!PctPcp = CDec(txtDato(11).Text)
      porstTGCfg!ImpUIT = CDec(txtDato(12).Text)
'      porstMain!CodMon = optMoneda(1).Value
   Else
      cboMonFnc.ListIndex = IIf(porstCOCfg!TpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      chkNiveles(0).Value = IIf(porstCOCfg!CodCta_Nv3 = 1, vbChecked, vbUnchecked)
      chkNiveles(1).Value = IIf(porstCOCfg!CodCta_Nv4 = 1, vbChecked, vbUnchecked)
      chkNiveles(2).Value = IIf(porstCOCfg!CodCta_Nv5 = 1, vbChecked, vbUnchecked)
      chkNiveles(3).Value = IIf(porstCOCfg!CodCta_Nv6 = 1, vbChecked, vbUnchecked)
      chkNiveles(4).Value = IIf(porstCOCfg!CodCta_Nv7 = 1, vbChecked, vbUnchecked)
      chkNiveles(5).Value = IIf(porstCOCfg!CodCta_Nv8 = 1, vbChecked, vbUnchecked)
'      dcoSocio.Item(0).BoundText = porstMain!CodSoc
'      dtpFecha.Value = porstMain!FehOpe
      optMon1.Value = IIf(porstCOCfg!IndMNE = INDMNE_INA, True, False)
      optMon2.Value = IIf(porstCOCfg!IndMNE = INDMNE_ACT, True, False)
'      optMoneda(1).Value = porstMain!CodMon
      txtDato(0).Text = IIf(IsNull(porstCOCfg!TpoMon_Sgn_MN), "", porstCOCfg!TpoMon_Sgn_MN)
      txtDato(1).Text = IIf(IsNull(porstCOCfg!TpoMon_Sgn_ME), "", porstCOCfg!TpoMon_Sgn_ME)
      txtDato(9).Text = IIf(IsNull(porstCOCfg!COdTDC_Pcp), "", porstCOCfg!COdTDC_Pcp)
      txtDato(6).Text = IIf(IsNull(porstCOCfg!CodTDc_Rtc), "", porstCOCfg!CodTDc_Rtc)
      txtDato(10).Text = IIf(IsNull(porstCOCfg!COdCta_Pcp), "", porstCOCfg!COdCta_Pcp)
      txtDato(7).Text = IIf(IsNull(porstCOCfg!CodCta_Rtc), "", porstCOCfg!CodCta_Rtc)
      txtDato(2).Text = Format(porstTGCfg!PctIGV, FORMATO_NUM_4)
      txtDato(3).Text = Format(porstTGCfg!PctISC, FORMATO_NUM_4)
      txtDato(4).Text = Format(porstTGCfg!PctIR4, FORMATO_NUM_4)
      txtDato(5).Text = Format(porstTGCfg!PctIES, FORMATO_NUM_4)
      txtDato(8).Text = Format(porstTGCfg!PctRtc, FORMATO_NUM_4)
      txtDato(11).Text = Format(porstTGCfg!PctPcp, FORMATO_NUM_4)
      txtDato(12).Text = Format(porstTGCfg!ImpUIT, FORMATO_NUM_4)
   End If
      
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

'Public Sub upDatosPredeterminados()    'Cambiar.
'   chkEstado.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
'End Sub

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   cboMonFnc.Enabled = tbHabilitar
   For dnContador = 0 To 5
      chkNiveles(dnContador).Enabled = tbHabilitar
   Next
   optMon1.Enabled = tbHabilitar
   optMon2.Enabled = tbHabilitar

  'Ayudas.
   cmdDatoAyud.Item(6).Enabled = tbHabilitar
   cmdDatoAyud.Item(7).Enabled = tbHabilitar
   cmdDatoAyud.Item(9).Enabled = tbHabilitar
   cmdDatoAyud.Item(10).Enabled = tbHabilitar
'   cmdDatoAyud(0).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
End Sub


Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 6, 9                           'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 7, 10                           'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

'[Código propio del formulario.

']

