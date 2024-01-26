VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPPDTEeFf 
   Caption         =   "[título]"
   ClientHeight    =   6105
   ClientLeft      =   510
   ClientTop       =   1455
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10440
   Begin VB.Frame frmParametro 
      Caption         =   " Parámetros "
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   7815
      TabIndex        =   32
      Top             =   3585
      Width           =   2535
      Begin VB.CheckBox chkCuenta 
         Caption         =   "Cuenta Equivalente"
         ForeColor       =   &H00C00000&
         Height          =   280
         Left            =   135
         TabIndex        =   33
         Top             =   420
         Width           =   2175
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "&Actualizar"
         Height          =   400
         Left            =   1350
         TabIndex        =   35
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdPlantilla 
         Caption         =   "&Plantilla"
         Height          =   400
         Left            =   105
         TabIndex        =   34
         Top             =   960
         Width           =   1100
      End
   End
   Begin VB.Frame frmUbicacion 
      Caption         =   " Carpeta "
      ForeColor       =   &H00FF0000&
      Height          =   2475
      Left            =   7815
      TabIndex        =   26
      Top             =   345
      Width           =   2535
      Begin VB.DirListBox dlbDirectorio 
         Height          =   1665
         Left            =   150
         TabIndex        =   29
         Top             =   690
         Width           =   2235
      End
      Begin VB.DriveListBox drvUnidad 
         Height          =   315
         Left            =   150
         TabIndex        =   28
         Top             =   400
         Width           =   2235
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Directorio :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   27
         Top             =   200
         Width           =   765
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   9075
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   3000
      Width           =   1260
   End
   Begin TabDlg.SSTab tabProceso 
      Height          =   4935
      Left            =   120
      TabIndex        =   24
      Top             =   105
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   494
      TabCaption(0)   =   "Estados Financieros"
      TabPicture(0)   =   "frmPPDTEeFf.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmRegistro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Balance de Comprobación"
      TabPicture(1)   =   "frmPPDTEeFf.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "mfgBalance"
      Tab(1).Control(1)=   "txtCelda"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtCelda 
         Enabled         =   0   'False
         Height          =   280
         Left            =   -74925
         TabIndex        =   40
         Top             =   4260
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Frame frmRegistro 
         Caption         =   " Registros "
         ForeColor       =   &H00FF0000&
         Height          =   3840
         Left            =   390
         TabIndex        =   0
         Top             =   525
         Width           =   5520
         Begin VB.TextBox txtDato 
            Height          =   280
            Index           =   5
            Left            =   180
            TabIndex        =   21
            Top             =   3405
            Width           =   585
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   5
            Left            =   5055
            Picture         =   "frmPPDTEeFf.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   3405
            Width           =   255
         End
         Begin VB.TextBox txtDato 
            Height          =   280
            Index           =   4
            Left            =   180
            TabIndex        =   17
            Top             =   2760
            Width           =   585
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   4
            Left            =   5055
            Picture         =   "frmPPDTEeFf.frx":01E2
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtDato 
            Height          =   280
            Index           =   3
            Left            =   180
            TabIndex        =   13
            Top             =   2115
            Width           =   585
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   3
            Left            =   5055
            Picture         =   "frmPPDTEeFf.frx":038C
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox txtDato 
            Height          =   280
            Index           =   2
            Left            =   180
            TabIndex        =   10
            Top             =   1770
            Width           =   585
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   2
            Left            =   5055
            Picture         =   "frmPPDTEeFf.frx":0536
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1770
            Width           =   255
         End
         Begin VB.TextBox txtDato 
            Height          =   280
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   1140
            Width           =   585
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   1
            Left            =   5055
            Picture         =   "frmPPDTEeFf.frx":06E0
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1140
            Width           =   255
         End
         Begin VB.TextBox txtDato 
            Height          =   280
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   525
            Width           =   585
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   0
            Left            =   5055
            Picture         =   "frmPPDTEeFf.frx":088A
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   525
            Width           =   255
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "Cuentas por pagar diversas"
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   20
            Top             =   3105
            Value           =   1  'Checked
            Width           =   3600
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "Cuentas por cobrar diversas"
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   9
            Top             =   1485
            Value           =   1  'Checked
            Width           =   3600
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "Proveedores"
            Height          =   240
            Index           =   3
            Left            =   180
            TabIndex        =   16
            Top             =   2475
            Value           =   1  'Checked
            Width           =   3600
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "Provisión ctas cobr. dudosa"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   855
            Value           =   1  'Checked
            Width           =   3600
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "Clientes"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   250
            Value           =   1  'Checked
            Width           =   3600
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   735
            TabIndex        =   22
            Top             =   3405
            Width           =   4305
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   735
            TabIndex        =   18
            Top             =   2760
            Width           =   4305
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   735
            TabIndex        =   14
            Top             =   2115
            Width           =   4305
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   735
            TabIndex        =   11
            Top             =   1770
            Width           =   4305
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   735
            TabIndex        =   7
            Top             =   1140
            Width           =   4305
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   735
            TabIndex        =   3
            Top             =   525
            Width           =   4305
         End
      End
      Begin MSFlexGridLib.MSFlexGrid mfgBalance 
         Bindings        =   "frmPPDTEeFf.frx":0A34
         Height          =   4365
         Left            =   -74835
         TabIndex        =   25
         Top             =   435
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   7699
         _Version        =   393216
         BackColorFixed  =   16777152
         ForeColorFixed  =   16711680
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   400
      Left            =   2340
      TabIndex        =   39
      Top             =   5655
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   400
      Left            =   4020
      TabIndex        =   38
      Top             =   5655
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   5340
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   180
      Index           =   1
      Left            =   8250
      TabIndex        =   30
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Archivo:"
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   5100
      Width           =   1785
   End
End
Attribute VB_Name = "frmPPDTEeFf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[Propio del formulario.
Private pocnnMain As ADODB.Connection
Private porstCOCta As ADODB.Recordset
Private porstBceCpb As ADODB.Recordset
Private s_Cuenta As String
Private a_Totales(13) As Double

Private Sub chkCuenta_Click()
  ppRegistrosGrilla
End Sub

Private Sub cmdAceptar_Click()
  On Error GoTo Err
  
  If MsgBox(Choose(gsIdioma, "Estás Seguro de Generar archivo de información ? ", " Are you sure you generate information file?"), vbQuestion + vbYesNo) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    cmdPlantilla.Enabled = False
    cmdActualiza.Enabled = False
    pgbProgreso(0).Value = 0: pgbProgreso(0).Min = 0
    
    ' Genero los archivos de información
    If tabProceso.Tab = 0 Then
      ppGenArchivoEeFf
    ElseIf tabProceso.Tab = 1 Then
      ppGenArchivoBceCpb
    End If
    
    MsgBox TEXT_8008, vbInformation
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    cmdPlantilla.Enabled = True
    cmdActualiza.Enabled = True
    cmdSalir.SetFocus
  End If
  Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
End Sub

Private Sub cmdActualiza_Click()
  On Error GoTo Err
    
  If MsgBox(Choose(gsIdioma, "Estás Seguro de Actualizar información ? ", "Are you sure  you update information ? "), vbQuestion + vbYesNo) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    cmdPlantilla.Enabled = False
    cmdActualiza.Enabled = False
    
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
     ' Recupera la información del balance
     ppGenBalance
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
    ' Recupera información importada
    ppRegistrosGrilla
    
    ' Mensaje de culminacion de proceso
    MsgBox TEXT_8008, vbInformation
    cmdAceptar.Enabled = True
    cmdPlantilla.Enabled = True
    cmdSalir.Enabled = True
    cmdActualiza.Enabled = True
    cmdSalir.SetFocus
  End If
  Exit Sub

Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdAceptar.Enabled = True
  cmdPlantilla.Enabled = True
  cmdActualiza.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  
  Select Case Index                   'Cambiar. Añadir índices.
  Case 0, 1, 2, 3, 4, 5
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index

End Sub

Private Sub cmdPlantilla_Click()
  On Error GoTo Err
    
  If MsgBox(Choose(gsIdioma, "Estás Seguro de Importar Plantilla de información ?", "Are you sure you import information attern?"), vbQuestion + vbYesNo) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    cmdPlantilla.Enabled = False
    cmdActualiza.Enabled = False
    
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
     ' Importa la plantilla
     ppImportoPlantilla
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
    ' Recupera información importada
    ppRegistrosGrilla
    
    ' Mensaje de culminacion de proceso
    MsgBox TEXT_8008, vbInformation
    lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Archivo:", "Processing File:")
    cmdAceptar.Enabled = True
    cmdPlantilla.Enabled = True
    cmdSalir.Enabled = True
    cmdActualiza.Enabled = True
    cmdSalir.SetFocus
  End If
  Exit Sub

Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdAceptar.Enabled = True
  cmdPlantilla.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub drvUnidad_Change()

dlbDirectorio.path = drvUnidad.Drive
dlbDirectorio.Refresh

End Sub

Private Sub Form_Activate()
   
  drvUnidad.Drive = gsRutSis
  dlbDirectorio.path = gsRutSis
  
  cmdSalir.SetFocus

End Sub
Private Sub Form_Load()
  Dim dnContador As Integer
  
  Set pocnnMain = New ADODB.Connection
  Set porstCOCta = New ADODB.Recordset
  Set porstBceCpb = New ADODB.Recordset
  
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND

  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  With porstCOCta
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
    .Source = .Source & "FROM COCta "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodCta"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  With porstBceCpb
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodCta, CodAux, DetCta, TpoCta, NumCol2, NumCol3, "
    .Source = .Source & "NumCol4, NumCol5, NumCol10, numCol11, NomRpt, codemp, pdoano "
    .Source = .Source & "FROM cotmprpt "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND NomRpt='tmpBceCpb' "
    .Source = .Source & "ORDER BY CodCta"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "cotmprpt"
  End With
 
 ']
  
  With txtDato
    For dnContador = 0 To 5
      .Item(dnContador).DataField = "CodCta"
      .Item(dnContador).MaxLength = 3
      .Item(dnContador).Text = Choose(dnContador + 1, "12", "19", "16", "422", "42", "46")
      ppAyuDet dnContador
    Next dnContador
  End With
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Directorio :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Directory :", "Currency :")
  Next nElemento
  tabProceso.TabCaption(0) = Choose(gsIdioma, "Estados Financieros", "Financial Statements")
  tabProceso.TabCaption(1) = Choose(gsIdioma, "Balance de Comprobación", "Trial Balance")
  frmRegistro.Caption = Choose(gsIdioma, " Registros ", " Registers ")
  chkInformacion(0).Caption = Choose(gsIdioma, "Clientes", "Customer")
  chkInformacion(1).Caption = Choose(gsIdioma, "Provision cuentas x cobrar dudosa", "Provision doubtful accounts receivable")
  chkInformacion(2).Caption = Choose(gsIdioma, "Cuentas por cobrar diversas", "Accounts receivable diverse")
  chkInformacion(3).Caption = Choose(gsIdioma, "Proveedores", "Suppliers")
  chkInformacion(4).Caption = Choose(gsIdioma, "Cuentas por pagar diversas", "Accounts payable diverse")
  frmUbicacion.Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  frmParametro.Caption = Choose(gsIdioma, " Parametros ", " Parameters ")
  chkCuenta.Caption = Choose(gsIdioma, "Cuenta Equivalente", "Equivalent Account")
  cmdPlantilla.Caption = Choose(gsIdioma, "&Plantilla", "&Pattern")
  cmdActualiza.Caption = Choose(gsIdioma, "&Actualizar", "&Update")
  lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Archivo:", "Processing File:")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
 
  ppRegistrosGrilla
  ' Configura el texto de la celda
  txtCelda.Text = 0
  txtCelda.Alignment = vbRightJustify
  txtCelda.MaxLength = 18
  txtCelda.Enabled = False
  txtCelda.Visible = False
  frmParametro.Visible = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
  porstBceCpb.Close
  porstCOCta.Close
  pocnnMain.Close
  Set porstBceCpb = Nothing
  Set porstCOCta = Nothing
  Set pocnnMain = Nothing
End Sub

Private Sub mfgBalance_DblClick()
  ' Verificar registros
  If mfgBalance.Rows = 3 Then Exit Sub
  If Not ((mfgBalance.Col >= 3 And mfgBalance.Col <= 6) Or (mfgBalance.Col >= 11 And mfgBalance.Col <= 12)) Then MsgBox Choose(gsIdioma, "Columna no actualizable", "This Col can not be up to date"), vbInformation: Exit Sub
  ppEditGrilla mfgBalance, txtCelda, vbKeySpace ' Simula un espacio
End Sub
Private Sub mfgBalance_KeyPress(KeyAscii As Integer)
  ' Verificar registros
  If mfgBalance.Rows = 3 Then Exit Sub
  If Not ((mfgBalance.Col >= 3 And mfgBalance.Col <= 6) Or (mfgBalance.Col >= 11 And mfgBalance.Col <= 12)) Then MsgBox Choose(gsIdioma, "Columna no actualizable", "This Col can not be up to date"), vbInformation: Exit Sub
  ppEditGrilla mfgBalance, txtCelda, KeyAscii
End Sub

Private Sub tabProceso_Click(PreviousTab As Integer)
  frmParametro.Visible = (tabProceso.Tab = 1)
End Sub

Private Sub txtCelda_GotFocus()
  txtCelda.SelStart = 0
  txtCelda.SelLength = txtCelda.MaxLength
End Sub
Private Sub txtCelda_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub txtCelda_KeyUp(KeyCode As Integer, Shift As Integer)
  ppCeldaGrilla
  ppEditKeyCode mfgBalance, txtCelda, KeyCode
End Sub
Private Sub txtCelda_LostFocus()
  txtCelda.Text = 0
  txtCelda.Enabled = False
  txtCelda.Visible = False
End Sub
Private Sub txtCelda_Validate(Cancel As Boolean)
  txtCelda.Text = IIf(Not IsNumeric(txtCelda.Text), 0, txtCelda.Text)
  txtCelda.Text = IIf(CDec(txtCelda.Text) < 0, 0, txtCelda.Text)
  txtCelda.Text = FormatNumber(txtCelda.Text, 0)
End Sub

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
  Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2, 3, 4, 5                       'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub

Private Function pfOutApoRet(s_Expresion As String) As String

s_Expresion = Trim$(s_Expresion)
If s_Expresion <> "" Then
    ' saco los enters de la cadena de caracteres
    While InStr(s_Expresion, Chr(13)) <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(13)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(13)) + 1)
    Wend
    ' saco los retornos de la cadena de caracteres
    While InStr(s_Expresion, Chr(10)) <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(10)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(10)) + 1)
    Wend
    ' saco los apostrofes de la cadena de caracteres
    While InStr(s_Expresion, "'") <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "'") - 1) & "´" & Mid$(s_Expresion, InStr(s_Expresion, "'") + 1)
    Wend
    ' saco los rayas de la cadena de caracteres
    While InStr(s_Expresion, "|") <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "|") - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, "|") + 1)
    Wend
End If
pfOutApoRet = Trim$(s_Expresion)

End Function

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0, 1, 2, 3, 4, 5                       'Cambiar (añadir índices).
    modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + frmRegistro.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + frmRegistro.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub
Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 0, 1, 2, 3, 4, 5
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstCOCta
      .MoveFirst
      .Find "CodCta='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
         MsgBox TEXT_8006, vbExclamation
         ppAyuDet = True
      Else
         lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
      End If
    End With
  End Select
End Function

Private Sub ppCeldaRecordset(n_Fila As Long, n_Columna As Long)
  Dim n_Importe As Double
  
  ' Ubico la cuenta a actualizar
  porstBceCpb.MoveFirst
  porstBceCpb.Find "CodCta='" & s_Cuenta & "'"
  If porstBceCpb.EOF Then
    MsgBox Choose(gsIdioma, "Celda no actualizable", "This Cell can not be up to date"), vbInformation
    Exit Sub
  End If
  n_Importe = CDec(porstBceCpb("numcol" & n_Columna - 1))
  ' Actualizo el importe de la cuenta
  pocnnMain.BeginTrans            'INICIA TRANSACCION.
  porstBceCpb("numcol" & n_Columna - 1) = CDec(mfgBalance.TextMatrix(n_Fila, n_Columna))
  porstBceCpb.Update
  pocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  ' Actualizo la fila de la grilla
  ppFilaGrilla n_Fila
  ' Actualiza el Total de la Columna
  a_Totales(n_Columna - 3) = Round(a_Totales(n_Columna - 3) - n_Importe, 0)
  n_Importe = CDec(mfgBalance.TextMatrix(n_Fila, n_Columna))
  a_Totales(n_Columna - 3) = Round(a_Totales(n_Columna - 3) + n_Importe, 0)
  mfgBalance.TextMatrix(1, n_Columna) = FormatNumber(a_Totales(n_Columna - 3), 0)
            
End Sub
Private Sub ppCeldaGrilla()
    
  If Not IsNumeric(txtCelda.Text) Then: MsgBox Choose(gsIdioma, "Ingresó Incorrecto, Los Valores deben ser numéricos", "Information isn't valid, The values must be numerics"), vbCritical, "Importe de Celda": Exit Sub
  If mfgBalance.Col >= 4 And mfgBalance.Col <= 12 Then
    mfgBalance.TextMatrix(mfgBalance.row, mfgBalance.Col) = FormatNumber(CDec(txtCelda.Text), 0)
  End If
  mfgBalance = FormatNumber(CDec(txtCelda.Text), 0)
  ppCeldaRecordset mfgBalance.row, mfgBalance.Col
  
End Sub
Private Sub ppEditGrilla(o_Grilla As MSFlexGrid, o_Texto As TextBox, n_KeyAscii As Integer)

  ' Utiliza el carácter escrito
  Select Case n_KeyAscii
   Case 0 To vbKeySpace   ' Un espacio significa modificar el texto actual
    o_Texto = o_Grilla
    o_Texto.SelStart = 1000
   Case Else              ' Otro carácter reemplaza el texto actual
    If Not IsNumeric(o_Texto.Text) Then: MsgBox Choose(gsIdioma, "Ingresó Incorrecto, Los Valores deben ser numéricos", "Information isn't valid, The values must be numerics"), vbCritical, "Importe de Celda": Exit Sub
    o_Texto.SelStart = 1
  End Select
  ' Muestra el celda en la posición correcta
  s_Cuenta = o_Grilla.TextMatrix(o_Grilla.row, 0)
  o_Texto.Move o_Grilla.Left + o_Grilla.CellLeft, o_Grilla.Top + o_Grilla.CellTop, o_Grilla.CellWidth, o_Grilla.CellHeight
  o_Texto.Visible = True
  o_Texto.Enabled = True
  o_Texto.SetFocus

End Sub
Private Sub ppEditKeyCode(o_Grilla As MSFlexGrid, o_Texto As TextBox, n_KeyCode As Integer)
    
  Select Case n_KeyCode
   Case vbKeyEscape, vbKeyReturn  ' ESC, ENTER: ocultar, devuelve el enfoque a la grilla
    o_Texto.Visible = False
    o_Grilla.SetFocus
   Case vbKeyUp                   ' Arriba: ocultar, devuelve el enfoque a la grilla
    o_Texto.Visible = False
    If o_Grilla.row > o_Grilla.FixedRows Then
        o_Grilla.row = o_Grilla.row - 1
    End If
    o_Grilla.SetFocus
   Case vbKeyDown                 ' Abajo: ocultar, devuelve el enfoque a la grilla
    o_Texto.Visible = False
    If o_Grilla.row < (o_Grilla.Rows - 1) Then
        o_Grilla.row = o_Grilla.row + 1
    End If
    o_Grilla.SetFocus
  End Select

End Sub
Private Sub ppFilaGrilla(n_Fila As Long)
  Dim n_aImporte(14) As Double, n_SalDeb As Double, n_SalHab As Double
  
  ' Obtengo los importes iniciales
  n_aImporte(1) = CDec(porstBceCpb!NumCol2)
  n_aImporte(2) = CDec(porstBceCpb!numCol3)
  n_aImporte(3) = CDec(porstBceCpb!numCol4)
  n_aImporte(4) = CDec(porstBceCpb!numCol5)
  ' Sumatorias y saldos
  n_aImporte(5) = Round(n_aImporte(1) + n_aImporte(3), 2)
  n_aImporte(6) = Round(n_aImporte(2) + n_aImporte(4), 2)
  n_aImporte(7) = Round(IIf(n_aImporte(5) >= n_aImporte(6), (n_aImporte(5) - n_aImporte(6)), 0), 2)
  n_aImporte(8) = Round(IIf(n_aImporte(6) >= n_aImporte(5), (n_aImporte(6) - n_aImporte(5)), 0), 2)
  ' Importes de transferencia
  n_aImporte(9) = CDec(porstBceCpb!numCol10)
  n_aImporte(10) = CDec(porstBceCpb!numCol11)
  ' Nuevos saldos
  n_SalDeb = Round(n_aImporte(7) + n_aImporte(9), 2)
  n_SalHab = Round(n_aImporte(8) + n_aImporte(10), 2)
  ' Cuentas balance, resultados
  n_aImporte(11) = Round(IIf(((n_SalDeb >= n_SalHab) And Left(porstBceCpb!CodCta, 1) <= "5"), (n_SalDeb - n_SalHab), 0), 2)
  n_aImporte(12) = Round(IIf(((n_SalHab >= n_SalDeb) And Left(porstBceCpb!CodCta, 1) <= "5"), (n_SalHab - n_SalDeb), 0), 2)
  n_aImporte(13) = Round(IIf(((n_SalDeb >= n_SalHab) And Left(porstBceCpb!CodCta, 1) >= "6"), (n_SalDeb - n_SalHab), 0), 2)
  n_aImporte(14) = Round(IIf(((n_SalHab >= n_SalDeb) And Left(porstBceCpb!CodCta, 1) >= "6"), (n_SalHab - n_SalDeb), 0), 2)
  
  ' Fila modificable de la grilla
  With mfgBalance
    .TextMatrix(n_Fila, 0) = porstBceCpb!CodCta
    .TextMatrix(n_Fila, 1) = IIf(IsNull(porstBceCpb!codaux), "", porstBceCpb!codaux)
    .TextMatrix(n_Fila, 2) = porstBceCpb!detcta
    .TextMatrix(n_Fila, 3) = FormatNumber(n_aImporte(1), 0)
    .TextMatrix(n_Fila, 4) = FormatNumber(n_aImporte(2), 0)
    .TextMatrix(n_Fila, 5) = FormatNumber(n_aImporte(3), 0)
    .TextMatrix(n_Fila, 6) = FormatNumber(n_aImporte(4), 0)
    .TextMatrix(n_Fila, 7) = FormatNumber(n_aImporte(5), 0)
    .TextMatrix(n_Fila, 8) = FormatNumber(n_aImporte(6), 0)
    .TextMatrix(n_Fila, 9) = FormatNumber(n_aImporte(7), 0)
    .TextMatrix(n_Fila, 10) = FormatNumber(n_aImporte(8), 0)
    .TextMatrix(n_Fila, 11) = FormatNumber(n_aImporte(9), 0)
    .TextMatrix(n_Fila, 12) = FormatNumber(n_aImporte(10), 0)
    .TextMatrix(n_Fila, 13) = FormatNumber(n_aImporte(11), 0)
    .TextMatrix(n_Fila, 14) = FormatNumber(n_aImporte(12), 0)
    .TextMatrix(n_Fila, 15) = FormatNumber(n_aImporte(13), 0)
    .TextMatrix(n_Fila, 16) = FormatNumber(n_aImporte(14), 0)
  End With

End Sub
Private Sub ppFilaTotal(ByVal n_FilaTotal As Integer)
  ' Fila de Totales Generales
  With mfgBalance
    .TextMatrix(n_FilaTotal, 0) = ""
    .TextMatrix(n_FilaTotal, 1) = ""
    .TextMatrix(n_FilaTotal, 2) = Choose(gsIdioma, "Totales Balance de Comprobación", "Totals Trial Balance")
    .TextMatrix(n_FilaTotal, 3) = FormatNumber(a_Totales(0), 0)
    .TextMatrix(n_FilaTotal, 4) = FormatNumber(a_Totales(1), 0)
    .TextMatrix(n_FilaTotal, 5) = FormatNumber(a_Totales(2), 0)
    .TextMatrix(n_FilaTotal, 6) = FormatNumber(a_Totales(3), 0)
    .TextMatrix(n_FilaTotal, 7) = FormatNumber(a_Totales(4), 0)
    .TextMatrix(n_FilaTotal, 8) = FormatNumber(a_Totales(5), 0)
    .TextMatrix(n_FilaTotal, 9) = FormatNumber(a_Totales(6), 0)
    .TextMatrix(n_FilaTotal, 10) = FormatNumber(a_Totales(7), 0)
    .TextMatrix(n_FilaTotal, 11) = FormatNumber(a_Totales(8), 0)
    .TextMatrix(n_FilaTotal, 12) = FormatNumber(a_Totales(9), 0)
    .TextMatrix(n_FilaTotal, 13) = FormatNumber(a_Totales(10), 0)
    .TextMatrix(n_FilaTotal, 14) = FormatNumber(a_Totales(11), 0)
    .TextMatrix(n_FilaTotal, 15) = FormatNumber(a_Totales(12), 0)
    .TextMatrix(n_FilaTotal, 16) = FormatNumber(a_Totales(13), 0)
  End With
End Sub
Private Sub ppGenArchivoBceCpb()
  
  Static sSentencia As String, sLinea As String
  Static nContador As Integer, nArchivo As Integer
  Static nRegistro As Double, nNumRegistros As Double
  Static sArchivo As String, sCaracter As String
  Static sImporte As String, n_SumCuenta As Double
  
  Static porstTmp As ADODB.Recordset

  ' Generacion de la tabla de seleccion
  sSentencia = "SELECT CodCta, CodAux, DetCta, TpoCta, NumCol2, NumCol3, "
  sSentencia = sSentencia & "NumCol4, NumCol5, NumCol10, NumCol11, NomRpt "
  sSentencia = sSentencia & "FROM cotmprpt "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND NomRpt='tmpBceCpb'"
  sSentencia = sSentencia & "ORDER BY CodCta"
  
  ' Seteo el recordset temporal
  Set porstTmp = New ADODB.Recordset
  sCaracter = "|"
  ' Obtengo el archivo de texto libre
  nArchivo = FreeFile
  
  ' Abro el recordset temporal
  With porstTmp
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  If Not (porstTmp.BOF And porstTmp.EOF) Then
    ' Barro todo el recordset y lo grabo en el archivo
    lblProgreso(0).Caption = Choose(gsIdioma, "Exportando Archivo: Balance de Comprobación", "Expoting File: Trial Balance")
    nNumRegistros = porstTmp.RecordCount
    pgbProgreso(0).Max = nNumRegistros
    pgbProgreso(0).Value = pgbProgreso(0).Min
    nRegistro = 0
    ' Nombre del archivo de texto a generar
    sArchivo = dlbDirectorio.path & "\" & "0702" & gsRUCEmp & gsAnoAct & ".txt"
    ' Elimino archivo de texto si existe
    If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
    If Dir$(sArchivo, vbNormal) = "" Then
      Open sArchivo For Output Access Write Lock Read Write As #nArchivo
      porstTmp.MoveFirst
      While Not porstTmp.EOF
        nRegistro = nRegistro + 1
        n_SumCuenta = Round((porstTmp!NumCol2 + porstTmp!numCol3 + porstTmp!numCol4 + porstTmp!numCol5 + porstTmp!numCol10 + porstTmp!numCol11), 0)
        nContador = porstTmp!TpoCTA
        If n_SumCuenta > 0 Then
          ' Diseño y grabo la linea en el archivo
          sLinea = ""
          sLinea = sLinea & porstTmp!CodCta & sCaracter
          ' Saldos iniciales - debe
          sImporte = Format(porstTmp!NumCol2, "############")
          sImporte = IIf(sImporte = "", IIf(nContador = 1, sImporte, ""), sImporte)
          sLinea = sLinea & sImporte & sCaracter
          ' Saldos iniciales - haber
          sImporte = Format(porstTmp!numCol3, "############")
          sImporte = IIf(sImporte = "", IIf(nContador = 1, sImporte, ""), sImporte)
          sLinea = sLinea & sImporte & sCaracter
          
          ' Movimientos del ejercicio - debe
          sImporte = Format(porstTmp!numCol4, "############")
          sImporte = IIf(sImporte = "", IIf(nContador = 1, sImporte, "0"), sImporte)
          sLinea = sLinea & sImporte & sCaracter
          ' Movimientos del ejercicio - debe
          sImporte = Format(porstTmp!numCol5, "############")
          sImporte = IIf(sImporte = "", IIf(nContador = 1, sImporte, "0"), sImporte)
          sLinea = sLinea & sImporte & sCaracter
          
          ' Transferencias y cancelaciones - debe
          sImporte = Format(porstTmp!numCol10, "############")
          sImporte = IIf(sImporte = "", IIf(nContador = 1, sImporte, "0"), sImporte)
          sLinea = sLinea & sImporte & sCaracter
          ' Transferencias y cancelaciones haber
          sImporte = Format(porstTmp!numCol11, "############")
          sImporte = IIf(sImporte = "", IIf(nContador = 1, sImporte, "0"), sImporte)
          sLinea = sLinea & sImporte & sCaracter
          'TC 29-02-2016
          sLinea = sLinea & "0" & sCaracter
          sLinea = sLinea & "0" & sCaracter
          
          ' Grabo la linea en el archivo
          Print #nArchivo, sLinea
        End If
        pgbProgreso(0).Value = nRegistro
        DoEvents
        porstTmp.MoveNext
      Wend
      ' cierro el archivo
      Close #nArchivo
    End If
    porstTmp.Close
  End If
  ' Cierro y saco de memoria los objetos
  Set porstTmp = Nothing

End Sub
Private Sub ppGenArchivoEeFf()
  
  Dim sSentencia As String, sLinea As String
  Dim nContador As Integer, nArchivo As Integer
  Dim nRegistro As Double, nNumRegistros As Double, nImporte As Double
  Dim sArchivo As String, nSecuencia As Integer
  Dim sCaracter  As String, sIndIngreso As String
  Dim nSumImporte As Double
  
  Dim porstTmp As ADODB.Recordset

  sSentencia = "SELECT LEFT(a.CodCta, 2) AS CodCta, a.CodAux AS CodAux, b.RazAux, b.RucAux, b.TpoDci, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpPdtEeFin ", "")
  sSentencia = sSentencia & "FROM (CocpbDet a "
  sSentencia = sSentencia & "LEFT JOIN TgAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<='13' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCta, '')<>'' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
  sSentencia = sSentencia & "AND LEFT(a.CodCta, 2) IN ("
  ' Cuentas a seleccionar
  If chkInformacion(0).Value = vbChecked Then
    sSentencia = sSentencia & "'" & Trim(txtDato(0).Text) & "', "
  End If
  If chkInformacion(1).Value = vbChecked Then
    sSentencia = sSentencia & "'" & Trim(txtDato(1).Text) & "', "
  End If
  If chkInformacion(3).Value = vbChecked Then
    sSentencia = sSentencia & "'" & Trim(txtDato(4).Text) & "', "
  End If
  If chkInformacion(4).Value = vbChecked Then
    sSentencia = sSentencia & "'" & Trim(txtDato(5).Text) & "'"
  End If
  sSentencia = sSentencia & ") "
  
  nContador = Len(Trim(txtDato(3).Text))
  sSentencia = sSentencia & "AND LEFT(a.CodCta, " & nContador & ")<>'" & Trim(txtDato(3).Text) & "' "
  sSentencia = sSentencia & "GROUP BY LEFT(a.CodCta, 2), a.CodAux, b.RazAux, b.RucAux, b.TpoDci "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 "
    sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) -  "
    sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  
  If chkInformacion(2).Value = vbChecked Then
    sSentencia = sSentencia & "UNION "
    sSentencia = sSentencia & "SELECT '" & Trim(txtDato(2).Text) & "' AS CodCta, a.CodAux AS CodAux, b.RazAux, b.RucAux, b.TpoDci, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
    sSentencia = sSentencia & "FROM (CocpbDet a "
    sSentencia = sSentencia & "LEFT JOIN TgAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND a.MesPvs<='13' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCta, '')<>'' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
    
    nContador = Len(Trim(txtDato(2).Text))
    sSentencia = sSentencia & "AND (LEFT(a.CodCta, " & nContador & ")='" & Trim(txtDato(2).Text) & "' "
    nContador = Len(Trim(txtDato(3).Text))
    sSentencia = sSentencia & "OR LEFT(a.CodCta, " & nContador & ")='" & Trim(txtDato(3).Text) & "') "
    
    sSentencia = sSentencia & "GROUP BY LEFT(a.CodCta, 2), a.CodAux, b.RazAux, b.RucAux, b.TpoDci "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
      sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 "
      sSentencia = sSentencia & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
      sSentencia = sSentencia & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
    End If
    sSentencia = sSentencia & "ORDER BY CodCta, CodAux"
  Else
    sSentencia = sSentencia & "ORDER BY a.CodCta, a.CodAux"
  End If
  ' Tabla temporal de saldos de cuentas
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpPdtEeFin", "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 13)='#tmpPdtEeFin_') DROP TABLE #tmpPdtEeFin")
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS tmpPdtEeFin ", "") & sSentencia
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Seteo el recordset temporal
  Set porstTmp = New ADODB.Recordset
  sCaracter = "|"
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkInformacion.Count - 1
    nSumImporte = 0
    ' Verifico que se haya seleccionado
    If chkInformacion(nContador).Value = vbChecked Then
      ' Obtengo el archivo de texto libre
      nArchivo = FreeFile
      ' Generacion de la tabla de seleccion
      sSentencia = "SELECT a.CodCta, a.CodAux, a.RazAux, a.RucAux, a.TpoDci, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.NomAux, '') AS NomAux, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.ApePatAux, '') AS ApePatAux, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.ApeMatAux, '') AS ApeMatAux, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(DebeSol-HaberSol), 0), 2) AS SaldoSol, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(DebeDol - HaberDol), 0), 2) As SaldoDol "
      sSentencia = sSentencia & "FROM (" & ps_Prefijo & "tmpPdtEeFin a "
      sSentencia = sSentencia & "LEFT JOIN tgauxnat b ON b.codemp='" & gsCodEmp & "' AND a.CodAux=b.CodAux) "
      sSentencia = sSentencia & "WHERE a.CodCta='" & Trim(txtDato(Choose(nContador + 1, 0, 1, 2, 4, 5)).Text) & "' "
      sSentencia = sSentencia & "GROUP BY CodCta, a.CodAux, a.RazAux, a.RucAux, a.TpoDci, b.NomAux, b.ApePatAux, b.ApeMatAux "
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & "HAVING (SaldoSol <> 0.00 Or SaldoDol <> 0.00) "
      ElseIf ps_Plataforma = pSrvSql Then
        sSentencia = sSentencia & "HAVING (ROUND(ISNULL(SUM(DebeSol-HaberSol), 0), 2) <> 0.00 "
        sSentencia = sSentencia & "OR ROUND(ISNULL(SUM(DebeDol - HaberDol), 0), 2) <> 0.00) "
      End If
      sSentencia = sSentencia & "ORDER BY a.CodCta, a.CodAux "
      ' Abro el recordset temporal
      With porstTmp
        If .State = adStateOpen Then .Close
        .ActiveConnection = pocnnMain
        .Source = sSentencia
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
      End With
      If Not (porstTmp.BOF And porstTmp.EOF) Then
        ' Barro todo el recordset y lo grabo en el archivo
        lblProgreso(0).Caption = Choose(gsIdioma, "Exportando Archivo: ", "Exporting File: ") & Mid(Trim(chkInformacion(nContador).Caption), 2)
        nNumRegistros = porstTmp.RecordCount
        pgbProgreso(0).Max = nNumRegistros
        pgbProgreso(0).Value = pgbProgreso(0).Min
        nRegistro = 0
        ' Nombre del archivo de texto a generar
        'sArchivo = dlbDirectorio.path & "\" & "0670" & gsRUCEmp & Choose(nContador + 1, "301", "303", "304", "335", "337") & ".txt"  &&tc
        'sArchivo = dlbDirectorio.path & "\" & "0682" & gsRUCEmp & Choose(nContador + 1, "361", "367", "364", "404", "407") & ".txt"
        sArchivo = dlbDirectorio.path & "\" & "0702" & gsRUCEmp & Choose(nContador + 1, "361", "367", "364", "404", "407") & ".txt"
        ' Elimino archivo de texto si existe
        If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
        If Dir$(sArchivo, vbNormal) = "" Then
          Open sArchivo For Output Access Write Lock Read Write As #nArchivo
          porstTmp.MoveFirst
          While Not porstTmp.EOF
            nImporte = Abs(IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, porstTmp!SaldoSol, porstTmp!SaldoDol))
            nSumImporte = Round(nSumImporte + nImporte, 2)
            nRegistro = nRegistro + 1
            ' Importe de acuerdo a parametro
            nImporte = Round(gnImpUIT * 3, 2)
            If Abs(porstTmp!SaldoSol) >= nImporte Then
              nImporte = Abs(IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, porstTmp!SaldoSol, porstTmp!SaldoDol))
              nSumImporte = Round(nSumImporte - nImporte, 2)
              ' Diseño y grabro la linea en el archivo
              sLinea = ""
              sLinea = sLinea & porstTmp!TpoDci & sCaracter
              sIndIngreso = "1"
              If Not (porstTmp!TpoDci = "00" Or porstTmp!TpoDci = "06") Then
                sLinea = sLinea & Right(porstTmp!rucaux, 8) & sCaracter
                sLinea = sLinea & sIndIngreso & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!ApePatAux <> "", porstTmp!ApePatAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!ApeMatAux <> "", porstTmp!ApeMatAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!NomAux <> "", porstTmp!NomAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!ApePatAux = "", porstTmp!razAux, "")) & sCaracter
              Else
                sIndIngreso = IIf(porstTmp!ApePatAux = "", "0", sIndIngreso)
                sLinea = sLinea & porstTmp!rucaux & sCaracter
                sLinea = sLinea & sIndIngreso & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(sIndIngreso = "1", porstTmp!ApePatAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(sIndIngreso = "1", porstTmp!ApeMatAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(sIndIngreso = "1", porstTmp!NomAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(sIndIngreso = "0", porstTmp!razAux, "")) & sCaracter
              End If
              sLinea = sLinea & Format(nImporte, "############") & sCaracter
              Print #nArchivo, sLinea
            End If
            pgbProgreso(0).Value = nRegistro
            DoEvents
            porstTmp.MoveNext
          Wend
          ' Diseño y grabro la linea de consolidados
          If nSumImporte > 0 Then
            sLinea = ""
            sLinea = sLinea & "99" & sCaracter
            sLinea = sLinea & "00000000000" & sCaracter
            sIndIngreso = "1"
            sLinea = sLinea & sIndIngreso & sCaracter
            sLinea = sLinea & sCaracter
            sLinea = sLinea & sCaracter
            sLinea = sLinea & sCaracter
            sLinea = sLinea & "CONSOLIDADO SALDOS MENORES A 3 UIT" & sCaracter
            sLinea = sLinea & Format(nSumImporte, "############") & sCaracter
            Print #nArchivo, sLinea
          End If
          ' cierro el archivo
          Close #nArchivo
        End If
      End If
      porstTmp.Close
    End If
  Next nContador
  ' Cierro y saco de memoria los objetos
  Set porstTmp = Nothing

End Sub
Private Sub ppGenBalance()
  
  Dim sSentencia As String, nLongitud As Integer
  Dim nContador As Integer, nNumRegistros As Long
  Dim nImporte As Double, nSumImporte As Double
  Dim s_SaldoDeb As String, s_SaldoHab As String
  Dim s_IniDebe As String, s_IniHaber As String
  Dim s_Moneda As String, s_Union As String
  
  ' Elimino temporal si existe
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 11)='#tmppdtbce_') DROP TABLE #tmppdtbce"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmppdtbce", sSentencia)
  
  ' Acumulación de saldos
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  s_SaldoDeb = "(": s_SaldoHab = "("
  For nContador = 1 To Val(gsMesAct)
    s_SaldoDeb = s_SaldoDeb & "acu.AcuD" & Format(Trim(nContador), "00") & "_" & s_Moneda & IIf(nContador = Val(gsMesAct), "", "+")
    s_SaldoHab = s_SaldoHab & "acu.AcuH" & Format(Trim(nContador), "00") & "_" & s_Moneda & IIf(nContador = Val(gsMesAct), "", "+")
  Next nContador
  s_SaldoDeb = s_SaldoDeb & ")"
  s_SaldoHab = s_SaldoHab & ")"
  s_IniDebe = "acu.AcuD00_" & s_Moneda
  s_IniHaber = "acu.AcuH00_" & s_Moneda
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmppdtbce ", "")
  ' Registro a dos digitos
  s_Union = "": nLongitud = 2
  For nContador = 2 To 5 ' de 4 a 5
    sSentencia = sSentencia & s_Union & "SELECT LEFT(acu.CodCta, " & nContador + nLongitud & ") AS CodCta, "
    If InStr(1, gsNivCta, Trim$(nContador)) > 0 Then
      sSentencia = sSentencia & "ROUND(" & s_IniDebe & ", 0) AS cSaldoD, "
      sSentencia = sSentencia & "ROUND(" & s_IniHaber & ", 0) AS cSaldoH, "
      sSentencia = sSentencia & "ROUND(" & s_SaldoDeb & ", 0) AS cSumaD, "
      sSentencia = sSentencia & "ROUND(" & s_SaldoHab & ", 0) AS cSumaH "
      sSentencia = sSentencia & IIf(s_Union = "" And ps_Plataforma = pSrvSql, "INTO #tmppdtbce ", "")
      sSentencia = sSentencia & "FROM CoCtaAcu acu "
      sSentencia = sSentencia & "WHERE acu.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND acu.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(acu.CodCta)=" & nContador & " "
    Else
      nLongitud = nContador + 1
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & s_IniDebe & ", 0)), 0) AS cSaldoD, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & s_IniHaber & ", 0)), 0) AS cSaldoH, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & s_SaldoDeb & ", 0)), 0) AS cSumaD, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & s_SaldoHab & ", 0)), 0) AS cSumaH "
      sSentencia = sSentencia & IIf(s_Union = "" And ps_Plataforma = pSrvSql, "INTO #tmppdtbce ", "")
      sSentencia = sSentencia & "FROM CoCtaAcu acu "
      sSentencia = sSentencia & "WHERE acu.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND acu.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(acu.CodCta)=" & nLongitud & " "
      sSentencia = sSentencia & "GROUP BY LEFT(acu.CodCta, " & nContador & ") "
    End If
    s_Union = "UNION "
    nLongitud = 0
  Next nContador
  sSentencia = sSentencia & "ORDER BY CodCta"
  ' genero tabla temporal con información del balance agrupado
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' Elimino temporal final si existe
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 11)='#rptpdtbce_') DROP TABLE #rptpdtbce"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS rptpdtbce", sSentencia)
  
  ' genero tabla temporal con saldos de plantilla
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE rptpdtbce ", "")
  sSentencia = sSentencia & "SELECT rpt.CodCta, rpt.CodAux, rpt.DetCta, rpt.TpoCta, rpt.Nomrpt, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(tmp.cSaldoD, 0.00) AS numCol2, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(tmp.cSaldoH, 0.00) AS numCol3, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(tmp.cSumaD, 0.00) AS numCol4, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(tmp.cSumaH, 0.00) AS numCol5 "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #rptpdtbce ", "")
  sSentencia = sSentencia & "FROM (cotmprpt rpt "
  If Not (chkCuenta.Value = vbChecked) Then
    sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmppdtbce tmp ON rpt.CodCta=tmp.CodCta) "
  Else
    sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmppdtbce tmp ON rpt.CodAux=tmp.CodCta) "
  End If
  sSentencia = sSentencia & "WHERE rpt.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND rpt.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND rpt.NomRpt='tmpBceCpb' "
  sSentencia = sSentencia & "ORDER BY rpt.CodCta "
  pocnnMain.Execute sSentencia, nNumRegistros
  
  ' paso final Elimino información del balance de comprobación
  sSentencia = "DELETE FROM cotmprpt WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND NomRpt='tmpBceCpb'"
  pocnnMain.Execute sSentencia, nNumRegistros
  ' paso final inserto información del balance de comprobación
  sSentencia = "INSERT INTO cotmprpt(codemp, pdoano, CodCta, CodAux, DetCta, TpoCta, numCol2, numCol3, numCol4, numCol5, NomRpt) "
  sSentencia = sSentencia & "SELECT '" & gsCodEmp & "', '" & gsAnoAct & "', CodCta, CodAux, DetCta, TpoCta, numCol2, numCol3, numCol4, numCol5, NomRpt "
  sSentencia = sSentencia & "FROM " & ps_Prefijo & "rptpdtbce "
  sSentencia = sSentencia & "ORDER BY CodCta"
  pocnnMain.Execute sSentencia, nNumRegistros

  ' Clase de cuenta 1
  ' Clase de cuenta 2
  ' Clase de cuenta 3
  
  ' Elimino las tablas temporales
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 11)='#rptpdtbce_') DROP TABLE #rptpdtbce"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS rptpdtbce", sSentencia)
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 11)='#tmppdtbce_') DROP TABLE #tmppdtbce"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmppdtbce", sSentencia)
  
End Sub
Private Sub ppImportoPlantilla()
    
  Dim sSentencia As String
  Dim nLongitud As Integer, nInicio As Integer
  Dim sArchivo As String, sMilinea As String
  Dim nContador As Integer, nArchivo As Integer
  Dim nRegistro As Long, nNumRegistros As Long
  
  ' Importo los registros
  nArchivo = FreeFile
  ' Abro Archivo de Texto
  sArchivo = dlbDirectorio.path & "\" & "pllbce" & ".txt"
  If Dir$(sArchivo, vbNormal) <> "" Then
    Open sArchivo For Input As #nArchivo
    nNumRegistros = CLng(LOF(nArchivo))
    If nNumRegistros > 0 Then
      ' Elimino los registros plantilla existente
      pocnnMain.Execute "DELETE FROM cotmprpt WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND nomrpt='tmpBceCpb'"
      ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
      ReDim aRegistros(4)
      lblProgreso(0).Caption = Choose(gsIdioma, "Importando Archivo: ", "Importing File: ") & sArchivo
      Do While Not EOF(nArchivo)
        Line Input #nArchivo, sMilinea
        nRegistro = nRegistro + 1
        nInicio = 1
        For nContador = 1 To 4
          nLongitud = Abs(InStr(nInicio, sMilinea, "|") - nInicio)
          aRegistros(nContador) = Mid$(sMilinea, nInicio, nLongitud)
          nInicio = nInicio + (nLongitud + 1)
        Next nContador
        ' Inserto registro
        sSentencia = "INSERT INTO cotmprpt(codemp, pdoano, CodCta, CodAux, DetCta, TpoCta, NomRpt) "
        sSentencia = sSentencia & "VALUES("
        sSentencia = sSentencia & "'" & gsCodEmp & "', "
        sSentencia = sSentencia & "'" & gsAnoAct & "', "
        sSentencia = sSentencia & "'" & aRegistros(1) & "', "
        If Not (aRegistros(2) = "") Then
          sSentencia = sSentencia & "'" & aRegistros(2) & "', "
        Else
          sSentencia = sSentencia & "Null, "
        End If
        sSentencia = sSentencia & "'" & aRegistros(3) & "', "
        sSentencia = sSentencia & "'" & aRegistros(4) & "', "
        sSentencia = sSentencia & "'tmpBceCpb')"
        pocnnMain.Execute sSentencia
      Loop
    End If
    Close #nArchivo
  End If

End Sub
Private Sub ppInicializaGrilla()
  Dim n_Index As Integer
    
  With mfgBalance
    .cols = 17
    .FixedCols = 3
    .Rows = 3
    .FixedRows = 2
    .GridColor = vbRed
    .GridColorFixed = vbBlue
    .Gridlines = flexGridFlat
    .GridLinesFixed = flexGridInset
    .GridLineWidth = 2
    .SelectionMode = flexSelectionFree
    .BackColor = &H80000018
    .BackColorBkg = &H8000000F
    .BackColorFixed = &HFFFFC0
    .BackColorSel = &H8000000D
    .ForeColor = vbBlack
    .ForeColorFixed = vbBlue
    .FillStyle = flexFillSingle
  End With
    
  For n_Index = 0 To (mfgBalance.cols - 1)
    mfgBalance.Col = n_Index
    If gsIdioma = NvlUsr_Sup Then
      mfgBalance.TextMatrix(0, n_Index) = Choose(n_Index + 1, "Cuenta", "Equival", "Descripción", "Ini Debe", "Ini Haber", "Mov Debe", "Mov Haber", "Sum Debe", "Sum Haber", "Sal Deudor", "Sal Acreedor", "Tra Debe", "Tra Haber", "Bal Activo", "Bal Pasivo", "Res Pérdidas", "Res Ganacias")
    Else
      mfgBalance.TextMatrix(0, n_Index) = Choose(n_Index + 1, "Account", "Equivalent", "Description", "Ini Debit", "Ini Credit", "Mov Debit", "Mov Credit", "Sum Debit", "Sum Credit", "Sal Debtor", "Sal Creditor", " Tra Debit", "Tra Credit", "Bal Activo", "Bal Pasivo", "Res Perdidas", "Res Ganancias")
    End If
    mfgBalance.ColAlignment(n_Index) = Choose(n_Index + 1, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter, flexAlignRightCenter)
    mfgBalance.ColWidth(n_Index) = Choose(n_Index + 1, 700, 700, 2700, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200)
  Next n_Index
  ' Fila de Totales
  ppFilaTotal 1
  ' Ancho de columna de cuenta equivalente
  mfgBalance.ColWidth(1) = IIf(chkCuenta.Value = vbChecked, 700, 0)
  ' Alto de fila adicional de totales
  mfgBalance.RowHeight(2) = 0
  
End Sub
Private Sub ppRegistrosGrilla()
  Dim n_Index As Integer
  Dim n_aImporte(14) As Double, n_SalDeb As Double, n_SalHab As Double
  
  mfgBalance.Redraw = False
  ' Elimino y configuro la grilla
  mfgBalance.Clear
  ppInicializaGrilla
  n_Index = 1
  a_Totales(0) = 0: a_Totales(1) = 0
  a_Totales(2) = 0: a_Totales(3) = 0
  a_Totales(4) = 0: a_Totales(5) = 0
  a_Totales(6) = 0: a_Totales(7) = 0
  a_Totales(8) = 0: a_Totales(9) = 0
  a_Totales(10) = 0: a_Totales(11) = 0
  a_Totales(12) = 0: a_Totales(13) = 0
  porstBceCpb.Requery
  If porstBceCpb.RecordCount > 0 Then
    porstBceCpb.MoveFirst
    n_Index = 3
    ' Fila de totales redusco altura a cero
    mfgBalance.RowHeight(n_Index - 1) = 0
    Do While Not porstBceCpb.EOF
      ' Obtengo los importes iniciales
      n_aImporte(1) = CDec(porstBceCpb!NumCol2)
      n_aImporte(2) = CDec(porstBceCpb!numCol3)
      n_aImporte(3) = CDec(porstBceCpb!numCol4)
      n_aImporte(4) = CDec(porstBceCpb!numCol5)
      ' Sumatorias y saldos
      n_aImporte(5) = Round(n_aImporte(1) + n_aImporte(3), 2)
      n_aImporte(6) = Round(n_aImporte(2) + n_aImporte(4), 2)
      n_aImporte(7) = Round(IIf(n_aImporte(5) >= n_aImporte(6), (n_aImporte(5) - n_aImporte(6)), 0), 2)
      n_aImporte(8) = Round(IIf(n_aImporte(6) >= n_aImporte(5), (n_aImporte(6) - n_aImporte(5)), 0), 2)
      ' Importes de transferencia
      n_aImporte(9) = CDec(porstBceCpb!numCol10)
      n_aImporte(10) = CDec(porstBceCpb!numCol11)
      ' Nuevos saldos
      n_SalDeb = Round(n_aImporte(7) + n_aImporte(9), 2)
      n_SalHab = Round(n_aImporte(8) + n_aImporte(10), 2)
      ' Cuentas balance, resultados
      n_aImporte(11) = Round(IIf(((n_SalDeb >= n_SalHab) And Left(porstBceCpb!CodCta, 1) <= "5"), (n_SalDeb - n_SalHab), 0), 2)
      n_aImporte(12) = Round(IIf(((n_SalHab >= n_SalDeb) And Left(porstBceCpb!CodCta, 1) <= "5"), (n_SalHab - n_SalDeb), 0), 2)
      n_aImporte(13) = Round(IIf(((n_SalDeb >= n_SalHab) And Left(porstBceCpb!CodCta, 1) >= "6"), (n_SalDeb - n_SalHab), 0), 2)
      n_aImporte(14) = Round(IIf(((n_SalHab >= n_SalDeb) And Left(porstBceCpb!CodCta, 1) >= "6"), (n_SalHab - n_SalDeb), 0), 2)
      With mfgBalance
        .Rows = n_Index
        .TextMatrix(n_Index - 1, 0) = porstBceCpb!CodCta
        .TextMatrix(n_Index - 1, 1) = IIf(IsNull(porstBceCpb!codaux), "", porstBceCpb!codaux)
        .TextMatrix(n_Index - 1, 2) = porstBceCpb!detcta
        .TextMatrix(n_Index - 1, 3) = FormatNumber(n_aImporte(1), 0)
        .TextMatrix(n_Index - 1, 4) = FormatNumber(n_aImporte(2), 0)
        .TextMatrix(n_Index - 1, 5) = FormatNumber(n_aImporte(3), 0)
        .TextMatrix(n_Index - 1, 6) = FormatNumber(n_aImporte(4), 0)
        .TextMatrix(n_Index - 1, 7) = FormatNumber(n_aImporte(5), 0)
        .TextMatrix(n_Index - 1, 8) = FormatNumber(n_aImporte(6), 0)
        .TextMatrix(n_Index - 1, 9) = FormatNumber(n_aImporte(7), 0)
        .TextMatrix(n_Index - 1, 10) = FormatNumber(n_aImporte(8), 0)
        .TextMatrix(n_Index - 1, 11) = FormatNumber(n_aImporte(9), 0)
        .TextMatrix(n_Index - 1, 12) = FormatNumber(n_aImporte(10), 0)
        .TextMatrix(n_Index - 1, 13) = FormatNumber(n_aImporte(11), 0)
        .TextMatrix(n_Index - 1, 14) = FormatNumber(n_aImporte(12), 0)
        .TextMatrix(n_Index - 1, 15) = FormatNumber(n_aImporte(13), 0)
        .TextMatrix(n_Index - 1, 16) = FormatNumber(n_aImporte(14), 0)
      End With
      ' Totalizo los importes
      a_Totales(0) = Round(a_Totales(0) + n_aImporte(1), 0)
      a_Totales(1) = Round(a_Totales(1) + n_aImporte(2), 0)
      a_Totales(2) = Round(a_Totales(2) + n_aImporte(3), 0)
      a_Totales(3) = Round(a_Totales(3) + n_aImporte(4), 0)
      a_Totales(4) = Round(a_Totales(4) + n_aImporte(5), 0)
      a_Totales(5) = Round(a_Totales(5) + n_aImporte(6), 0)
      a_Totales(6) = Round(a_Totales(6) + n_aImporte(7), 0)
      a_Totales(7) = Round(a_Totales(7) + n_aImporte(8), 0)
      a_Totales(8) = Round(a_Totales(8) + n_aImporte(9), 0)
      a_Totales(9) = Round(a_Totales(9) + n_aImporte(10), 0)
      a_Totales(10) = Round(a_Totales(10) + n_aImporte(11), 0)
      a_Totales(11) = Round(a_Totales(11) + n_aImporte(12), 0)
      a_Totales(12) = Round(a_Totales(12) + n_aImporte(13), 0)
      a_Totales(13) = Round(a_Totales(13) + n_aImporte(14), 0)
      
      ' Incremento las filas
      n_Index = n_Index + 1
      porstBceCpb.MoveNext
    Loop
  End If
  ' Visualizo, Inmovilizo los Totales
  ppFilaTotal 1
  If mfgBalance.Rows = 3 Then
    mfgBalance.FixedRows = IIf(mfgBalance.Rows = 0, 3, 2)
    mfgBalance.RowHeight(2) = IIf(mfgBalance.Rows = 0, 0, 240)
    mfgBalance.Redraw = True
  Else
    mfgBalance.FixedRows = IIf(mfgBalance.Rows = 3, 3, 2)
    mfgBalance.RowHeight(2) = IIf(mfgBalance.Rows = 3, 0, 240)
    mfgBalance.Redraw = True
  End If
End Sub
