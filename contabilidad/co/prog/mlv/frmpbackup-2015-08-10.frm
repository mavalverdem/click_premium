VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPBackup 
   Caption         =   "[título]"
   ClientHeight    =   7785
   ClientLeft      =   1815
   ClientTop       =   1245
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8355
   Begin VB.CommandButton cmdMysqlSql 
      Cancel          =   -1  'True
      Caption         =   "&Mysl a sql"
      Default         =   -1  'True
      Height          =   350
      Left            =   6840
      TabIndex        =   49
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab tabProceso 
      Height          =   6135
      Left            =   120
      TabIndex        =   41
      Top             =   60
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   494
      TabCaption(0)   =   "Configuración de Parámetros"
      TabPicture(0)   =   "frmpbackup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmTablas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmTransacciones"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmUbicacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmProceso"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame frmProceso 
         Caption         =   " Tipo de Proceso "
         ForeColor       =   &H00000080&
         Height          =   1125
         Left            =   5505
         TabIndex        =   37
         Top             =   2820
         Width           =   2535
         Begin VB.CheckBox chkProceso 
            Caption         =   "Restore General"
            ForeColor       =   &H00800000&
            Height          =   200
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.OptionButton optProceso 
            Caption         =   "&Restore de Información"
            ForeColor       =   &H00C00000&
            Height          =   200
            Index           =   1
            Left            =   240
            TabIndex        =   39
            Top             =   460
            Width           =   2000
         End
         Begin VB.OptionButton optProceso 
            Caption         =   "&Backup de Información"
            ForeColor       =   &H00C00000&
            Height          =   200
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   2000
         End
      End
      Begin VB.Frame frmUbicacion 
         Caption         =   " Carpeta "
         ForeColor       =   &H00000080&
         Height          =   2325
         Left            =   5505
         TabIndex        =   33
         Top             =   350
         Width           =   2535
         Begin VB.DriveListBox drvUnidad 
            Height          =   315
            Left            =   150
            TabIndex        =   35
            Top             =   400
            Width           =   2235
         End
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Left            =   150
            TabIndex        =   36
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   34
            Top             =   200
            Width           =   765
         End
      End
      Begin VB.Frame frmTransacciones 
         Caption         =   " Transacciones "
         ForeColor       =   &H00000080&
         Height          =   2025
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   7920
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Se&rvicios de Ventas"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   17
            Top             =   525
            Value           =   1  'Checked
            Width           =   2350
         End
         Begin VB.CheckBox checkincluir 
            Caption         =   "Incluir"
            Height          =   195
            Left            =   6015
            TabIndex        =   30
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox Check 
            Caption         =   "Seleccion"
            Height          =   195
            Left            =   4815
            TabIndex        =   29
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtDato 
            Enabled         =   0   'False
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
            Left            =   4815
            TabIndex        =   31
            Top             =   1560
            Width           =   525
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   7455
            Picture         =   "frmpbackup.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1560
            Width           =   255
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "&Saldos de Cuentas"
            Height          =   240
            Index           =   7
            Left            =   2745
            TabIndex        =   23
            Top             =   525
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Comprobantes de &Diario"
            Height          =   240
            Index           =   6
            Left            =   2745
            TabIndex        =   22
            Top             =   270
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Caja &Bancos"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   21
            Top             =   1560
            Value           =   1  'Checked
            Width           =   2350
         End
         Begin VB.ComboBox cmbPeriodo 
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   1
            Left            =   5625
            TabIndex        =   28
            Text            =   "Final"
            Top             =   840
            Width           =   2000
         End
         Begin VB.ComboBox cmbPeriodo 
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   0
            Left            =   5625
            TabIndex        =   26
            Text            =   "Inicio"
            Top             =   480
            Width           =   2000
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "&Pedidos de Compras"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   16
            Top             =   270
            Value           =   1  'Checked
            Width           =   2350
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Registro de &Compras"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   18
            Top             =   780
            Value           =   1  'Checked
            Width           =   2350
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Registro de &Ventas"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   19
            Top             =   1050
            Value           =   1  'Checked
            Width           =   2350
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Registro de &Honorarios"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   20
            Top             =   1305
            Value           =   1  'Checked
            Width           =   2350
         End
         Begin VB.Label lblTexto 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fin :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   5070
            TabIndex        =   27
            Top             =   885
            Width           =   465
         End
         Begin VB.Label lblTexto 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   5070
            TabIndex        =   25
            Top             =   525
            Width           =   465
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rango de periodos :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   5070
            TabIndex        =   24
            Top             =   210
            Width           =   1440
         End
         Begin VB.Shape shpCuadro 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C00000&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00FFC0C0&
            Height          =   1095
            Left            =   4830
            Shape           =   4  'Rounded Rectangle
            Top             =   165
            Width           =   3015
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
            Left            =   5295
            TabIndex        =   32
            Top             =   1560
            Width           =   2115
         End
      End
      Begin VB.Frame frmTablas 
         Caption         =   " Tablas "
         ForeColor       =   &H00000080&
         Height          =   3600
         Left            =   120
         TabIndex        =   0
         Top             =   350
         Width           =   5310
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Otros"
            Height          =   240
            Index           =   13
            Left            =   2790
            TabIndex        =   14
            Top             =   525
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "EE &Financieros Presupuesto  "
            Height          =   240
            Index           =   12
            Left            =   2790
            TabIndex        =   13
            Top             =   270
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Bancos"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   2
            Top             =   525
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "A&siento Tipo"
            Height          =   240
            Index           =   11
            Left            =   150
            TabIndex        =   12
            Top             =   3120
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Configuración &Reportes"
            Height          =   240
            Index           =   10
            Left            =   150
            TabIndex        =   11
            Top             =   2865
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Proyecto"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   10
            Top             =   2610
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Ta&bla de Configuración"
            Height          =   240
            Index           =   8
            Left            =   150
            TabIndex        =   9
            Top             =   2340
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Tipo de Ca&mbio"
            Height          =   240
            Index           =   7
            Left            =   150
            TabIndex        =   8
            Top             =   2085
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Flujo de Caja/Efectivo"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   7
            Top             =   1830
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Centro de Costo"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   6
            Top             =   1560
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Tipo de Documentos"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   5
            Top             =   1305
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Auxiliares"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   4
            Top             =   1050
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Plan de Cuentas"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   3
            Top             =   780
            Value           =   1  'Checked
            Width           =   2355
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Diario"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   270
            Value           =   1  'Checked
            Width           =   2355
         End
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   350
      Left            =   2685
      TabIndex        =   47
      Top             =   7350
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   350
      Left            =   4365
      TabIndex        =   46
      Top             =   7350
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   6510
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   45
      Top             =   6990
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
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
      Index           =   1
      Left            =   120
      TabIndex        =   44
      Top             =   6750
      Width           =   1785
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Información ..."
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
      TabIndex        =   42
      Top             =   6270
      Width           =   2310
   End
End
Attribute VB_Name = "frmPBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pocnnMain As ADODB.Connection
Public porstCodro As ADODB.Recordset
Private Sub Check_Click()
If Check.Value = Checked Then
    cmdDatoAyud(0).Enabled = True
Else
    cmdDatoAyud(0).Enabled = False
End If
End Sub
Private Sub chkImporProceso_Click(Index As Integer)
If Index = 5 Then
    If chkImporProceso(5).Value = Checked Then
        Check.Enabled = True
    Else
        Check.Enabled = False
    End If
End If
End Sub
Private Sub chkProceso_Click()
Static nContador As Integer

    If chkProceso.Value = vbChecked Then
        For nContador = 0 To chkImporTabla.Count - 1
            chkImporTabla(nContador).Value = vbChecked
            If nContador < 7 Then chkImporProceso(nContador).Value = vbChecked
        Next nContador
    End If
    frmTablas.Enabled = (chkProceso.Value = vbUnchecked)
    frmTransacciones.Enabled = (chkProceso.Value = vbUnchecked)

End Sub

Private Sub cmdAceptar_Click()
 
 '   On Error GoTo Err
    
  Dim sMensaje As String
  Dim plEliminar As Boolean
    
  plEliminar = True
  sMensaje = IIf(optProceso(0).Value, Choose(gsIdioma, "Generar archivo de Backup de información ?", "Do you want To Generate Backup file? "), Choose(gsIdioma, " Eliminar información existente de las tablas seleccionadas y las Transacciones por periodos; Restaurar la información del archivo de Backup ?", " Do you want to eliminate existing information in selected tables and transactions for periods; Restore Backup file ?"))
  If cmbPeriodo(0).ListIndex > cmbPeriodo(1).ListIndex Then Beep: MsgBox Choose(gsIdioma, "El Periodo de Inicio debe ser menor o igual al Periodo Final", " The beginning period  must be less or equal than End period"), vbExclamation: cmbPeriodo(0).SetFocus: Exit Sub
  If MsgBox(Choose(gsIdioma, "¿ Estás Seguro de ", "Are you sure ") & sMensaje, vbQuestion + vbYesNo) = vbYes Then
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
    pgbProgreso(0).Value = 0: pgbProgreso(0).Min = 0
    pgbProgreso(1).Value = 0: pgbProgreso(1).Min = 0
   
    ' Seteo y activo la coneccion
    Set pocnnMain = New ADODB.Connection
    With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .CommandTimeout = 180
      .Open
    End With
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
    If optProceso(0).Value Then
      ' Paso 1: Realizo la exportacion de las tablas
      pgbProgreso(0).Max = 2
      pgbProgreso(0).Value = pgbProgreso(0).Min
      ppBackup_Tablas
      pgbProgreso(0).Value = 1                    ' Actualizo la barra de progreso
      ' Paso 2 : Realizo la exportacion de las tablas de transacciones
      ppBackup_Proceso
      pgbProgreso(0).Value = 2                    ' Actualizo la barra de progreso
    Else
      ' Paso 1: Realizo la transformacion de informacion de las tablas
      pgbProgreso(0).Max = 5
      pgbProgreso(0).Value = pgbProgreso(0).Min
      ppRestore_Tablas
      pgbProgreso(0).Value = 1                    ' Actualizo la barra de progreso
      ' Paso 2 : Realizo la transformacion de informacion de las transacciones
      ppRestore_Proceso
      pgbProgreso(0).Value = 2                    ' Actualizo la barra de progreso
      ' Paso 3 : Realizo la verificacion de limpieza de tablas y procesos
      plEliminar = pfDel_Registros
      If Not plEliminar Then GoTo Err
      pgbProgreso(0).Value = 3                    ' Actualizo la barra de progreso
      ' Paso 4 : Realizo la eliminacion(validacion) y restauracion de tablas
      ppValida_Tablas
      pgbProgreso(0).Value = 4                    ' Actualizo la barra de progreso
      ' Paso 5 : Realizo la eliminacion(validacion) y restauracion de transacciones
      ppValida_Proceso
      pgbProgreso(0).Value = 5                    ' Actualizo la barra de progreso
    End If
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
  
    MsgBox TEXT_8008, vbInformation
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    cmdSalir.SetFocus
    pocnnMain.Close
    Set pocnnMain = Nothing
  End If
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  If plEliminar Then MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  
   ppAyuBus Index
End Sub

Private Sub cmdMysqlSql_Click()
      frmpbackup_strucompa.Show vbModal
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Function pfDel_Registros() As Boolean
  Dim sSentencia As String, sMilinea As String
  Dim nContador As Integer, nSecuencia As Integer
  Dim nNumRegistros As Long
  Dim aTabla(), aWhere(), aCampos(2, 2) As String
  Dim sWhereIni As String
    
  ' Obtengo los campos a actualizar si no es general
  If chkProceso.Value = vbUnchecked Then
    For nContador = cmbPeriodo(0).ListIndex To cmbPeriodo(1).ListIndex
      aCampos(1, 1) = aCampos(1, 1) & "AcuD" & Format(nContador, "00") & "_MN=0.00, "
      aCampos(1, 2) = aCampos(1, 2) & "AcuH" & Format(nContador, "00") & "_MN=0.00, "
      aCampos(2, 1) = aCampos(2, 1) & "AcuD" & Format(nContador, "00") & "_ME=0.00, "
      aCampos(2, 2) = aCampos(2, 2) & "AcuH" & Format(nContador, "00") & "_ME=0.00" & IIf(nContador = cmbPeriodo(1).ListIndex, "", ", ")
    Next nContador
  End If
   
  ' Elimino los registros de las transacciones
  For nContador = chkImporProceso.Count - 1 To 0 Step -1
    If chkImporProceso(nContador).Value = vbChecked Then
      aTabla = Choose(nContador + 1, Array("copdocpr", "copdocprcta"), Array("coconser"), Array("CoCprDoc", "CoCprDocCta", "CoCprDocCCo"), Array("CoVtaDoc", "CoVtaDocCta", "CoVtaDocCCo"), Array("CoHprDoc", "CoHprDocCta", "CoHprDocCCo"), Array("CoBanCab", "CoBanDet"), Array("CoCpbCab", "CoCpbDet", "CoCpbDetRP", "CoCpbDetFjo"), Array("coctaacu", "coauxacu", "coccoacu"))
      aWhere = Choose(nContador + 1, Array("", ""), Array(""), Array("", "AND a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc", "AND a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc"), Array("", "AND a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc", "AND a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc"), Array("", "AND a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc", "AND a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc"), Array("", ""), Array("", "", "", ""), Array("", "", ""))
      For nSecuencia = UBound(aTabla, 1) To 0 Step -1
        sMilinea = IIf(nSecuencia = 0 Or (nContador = 5 Or nContador = 6), "a", "b") & "."
        sSentencia = "DELETE a "
        sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " a "
        sWhereIni = "WHERE a.codemp='" & gsCodEmp & "' "
        sWhereIni = sWhereIni & "AND a.pdoano='" & gsAnoAct & "'"
        
        
        'sWhereIni = sWhereIni & "AND a.pdoano='" & gsAnoAct & "' and a.coddro <> '5001' "
        'If (aTabla(nSecuencia) = "CoCpbCab" Or aTabla(nSecuencia) = "CoCpbDet") And Check.Value = 1 Then
        '        If checkincluir.Value = 1 Then
        '            sWhereIni = sWhereIni & " and a.coddro in ('" & txtDato(0).Text & "') "
        '        Else
        '            sWhereIni = sWhereIni & " and a.coddro not in ('" & txtDato(0).Text & "') "
        '        End If
        'End If
        
        
        If nContador <> 7 Then
          ' Verifico existan registros del ejercicio
          pocnnMain.Execute "SELECT DISTINCTROW a.codemp, a.pdoano FROM " & ps_Prefijo & "tmp" & aTabla(0) & " a " & sWhereIni, nNumRegistros
          If nNumRegistros = 0 Then GoTo ManejaError
          
          If chkProceso.Value = vbUnchecked Then
            sSentencia = sSentencia & IIf(nSecuencia = 0 Or (nContador = 5 Or nContador = 6), "", ", " & aTabla(0) & " b ")
            sWhereIni = sWhereIni & "AND " & sMilinea & "MesPvs>='" & gfCeros(cmbPeriodo(0).ListIndex, 2, 0, "0") & "' "
            sWhereIni = sWhereIni & "AND " & sMilinea & "MesPvs<='" & gfCeros(cmbPeriodo(1).ListIndex, 2, 0, "0") & "' "
            sWhereIni = sWhereIni & IIf(nSecuencia = 0, "", aWhere(nSecuencia))
          End If
        Else
          If chkProceso.Value = vbUnchecked Then
            sSentencia = "UPDATE " & aTabla(nSecuencia) & " SET "
            sSentencia = sSentencia & aCampos(1, 1) & aCampos(1, 2)
            sSentencia = sSentencia & aCampos(2, 1) & aCampos(2, 2) & " "
            sWhereIni = "WHERE codemp='" & gsCodEmp & "' "
            sWhereIni = sWhereIni & "AND pdoano='" & gsAnoAct & "'"
          End If
        End If
        sSentencia = sSentencia & sWhereIni
        pocnnMain.Execute sSentencia, nNumRegistros
        DoEvents
      Next nSecuencia
    End If
  Next nContador
    
  ' Elimino los registros de las tablas
  For nContador = chkImporTabla.Count - 1 To 0 Step -1
    ' Verifico que se haya seleccionado
    If chkImporTabla(nContador).Value = vbChecked Then
      aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("codpe"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
      aWhere = Choose(nContador + 1, Array("S"), Array("N"), Array("S"), Array("N", "N"), Array("N"), Array("S"), Array("S", "S"), Array("N", "S"), Array("S", "S"), Array("S", "S", "S"), Array("S", "S", "S", "S"), Array("S", "S"), Array("N"), Array("N", "N", "N", "N"))
      For nSecuencia = UBound(aTabla, 1) To 0 Step -1
        ' Verifico si es por años
        If aWhere(nSecuencia) = "S" Then
          sSentencia = "DELETE a "
          sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " a "
          sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
          sSentencia = sSentencia & IIf(aWhere(nSecuencia) = "S", "AND a.pdoano='" & gsAnoAct & "'", "")
          pocnnMain.Execute sSentencia, nNumRegistros
        End If
        DoEvents
      Next nSecuencia
    End If
  Next nContador
  pfDel_Registros = True
  Exit Function

ManejaError:
  MsgBox Choose(gsIdioma, "Información a procesar no es del ejercicio activo", "Information to process is not of the active exercise"), vbExclamation
End Function

Private Function pfRegistro_Texto(ByVal rsRegistro As ADODB.Recordset, ByVal nColumnas As Integer, bTipo As Byte) As String
    Static nCampo As Integer, sTipoDato As String
    
    pfRegistro_Texto = ""
    For nCampo = 0 To nColumnas - 1
      sTipoDato = IIf((rsRegistro(nCampo).Type = adSmallInt Or rsRegistro(nCampo).Type = adInteger Or rsRegistro(nCampo).Type = adDouble Or rsRegistro(nCampo).Type = adCurrency Or rsRegistro(nCampo).Type = adNumeric), "N", IIf(rsRegistro(nCampo).Type = adChar Or rsRegistro(nCampo).Type = adVarChar, "C", IIf(rsRegistro(nCampo).Type = adDBDate, "F", IIf(rsRegistro(nCampo).Type = adDBTimeStamp, "T", "C"))))
      If (sTipoDato = "T" Or sTipoDato = "F") And bTipo = 2 Then
        pfRegistro_Texto = pfRegistro_Texto & pfSacaApoRet(Trim$(IIf(IsNull(rsRegistro(nCampo).Value), "", IIf(sTipoDato = "F", Format(rsRegistro(nCampo).Value, "dd/mm/yyyy"), Format(rsRegistro(nCampo).Value, "dd/mm/yyyy hh:mm:ss"))))) & "|"
      Else
        pfRegistro_Texto = pfRegistro_Texto & pfSacaApoRet(Trim$(Choose(bTipo + 1, rsRegistro(nCampo).Name, sTipoDato, IIf(IsNull(rsRegistro(nCampo).Value), "", rsRegistro(nCampo).Value)))) & "|"
      End If
    Next nCampo

End Function

Private Function pfSacaApoRet(s_Expresion As String) As String

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
pfSacaApoRet = Trim$(s_Expresion)

End Function

Private Sub ppBackup_Proceso()
  Dim sSentencia As String, sMilinea As String
  Dim nContador As Integer, nArchivo As Integer
  Dim nRegistro As Double, nNumRegistros As Double
  Dim sArchivo As String, nSecuencia As Integer
  Dim aTabla(), aArchivo(), aColumnas()
  Dim sWhere As String, sWhereIni As String
  Dim porstTmp As ADODB.Recordset

    ' Seteo el recordset temporal para la grabacion
    Set porstTmp = New ADODB.Recordset
    ' Importo las tablas de acuerdo a la selección
    For nContador = 0 To chkImporProceso.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporProceso(nContador).Value Then
        ' Obtengo el archivo de texto
        aArchivo = Choose(nContador + 1, Array("pdo", "pdc"), Array("svc"), Array("cpd", "cpc", "cpo"), Array("vtd", "vtc", "vto"), Array("hrd", "hrc", "hro"), Array("cbc", "cbd"), Array("cdc", "cdd", "cdr", "cdf"), Array("sct", "sax", "sco"))
        '2014-10-17 aumento cab pedidoa Columnas = Choose(nContador + 1, Array(21, 14), Array(16), Array(88, 18, 16), Array(75, 19, 15), Array(45, 17, 15), Array(28, 26), Array(15, 37, 24, 16), Array(63, 64, 64))
        '2015-08-04 verificar estructura de la base de datos  aColumnas = Choose(nContador + 1, Array(22, 14), Array(16), Array(88, 18, 16), Array(75, 19, 15), Array(45, 17, 15), Array(28, 26), Array(15, 37, 24, 16), Array(63, 64, 64))
        ' Genero el archivo y abro el recordset temporal
        aTabla = Choose(nContador + 1, Array("copdocpr", "copdocprcta"), Array("coconser"), Array("CoCprDoc", "CoCprDocCta", "CoCprDocCCo"), Array("CoVtaDoc", "CoVtaDocCta", "CoVtaDocCCo"), Array("CoHprDoc", "CoHprDocCta", "CoHprDocCCo"), Array("CoBanCab", "CoBanDet"), Array("CoCpbCab", "CoCpbDet", "CoCpbDetRP", "CoCpbDetFjo"), Array("CoCtaAcu", "CoAuxAcu", "CoCCoAcu"))
        sWhere = Choose(nContador + 1, "AND b.codemp=a.codemp AND b.pdoano=a.pdoano AND b.mespvs=a.mespvs AND b.coddpe=a.coddpe AND b.pdocpr=a.pdocpr", "AND b.codemp=a.codemp AND b.pdoano=a.pdoano AND b.mespvs=a.mespvs AND b.codcon=a.codcon", "AND b.codemp=a.codemp AND b.pdoano=a.pdoano AND b.CodAux=a.CodAux AND b.CodTDc=a.CodTDc AND b.SerDoc=a.SerDoc AND b.NroDoc=a.NroDoc", "AND b.codemp=a.codemp AND b.pdoano=a.pdoano AND b.CodTDc=a.CodTDc AND b.SerDoc=a.SerDoc AND b.NroDoc=a.NroDoc", "AND b.codemp=a.codemp AND b.pdoano=a.pdoano AND b.CodAux=a.CodAux AND b.SerDoc=a.SerDoc AND b.NroDoc=a.NroDoc", "", "", "")
        
'''ini 2015-08-04 verificar estructura de la base de datos
        aColumnas = ftot_field_arr(aTabla)
'''fin 2015-08-04 verificar estructura de la base de datos
        
        For nSecuencia = 0 To UBound(aArchivo, 1)
          ' Obtengo el archivo de texto libre
          nArchivo = FreeFile
          sMilinea = IIf(nSecuencia = 0 Or nContador = 5 Or nContador = 6, "a", "b")
          sSentencia = "SELECT a.* "
          sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " a "
          sWhereIni = "WHERE a.codemp='" & gsCodEmp & "' AND a.pdoano='" & gsAnoAct & "' "
          If nContador <> 7 Then
            sSentencia = sSentencia & IIf(nSecuencia = 0 Or nContador = 5 Or nContador = 6, "", ", " & aTabla(0) & " b ")
            sWhereIni = sWhereIni & "AND " & sMilinea & ".MesPvs>='" & gfCeros(cmbPeriodo(0).ListIndex, 2, 0, "0") & "' "
            sWhereIni = sWhereIni & "AND " & sMilinea & ".MesPvs<='" & gfCeros(cmbPeriodo(1).ListIndex, 2, 0, "0") & "' "
            sWhereIni = sWhereIni & IIf(nSecuencia = 0, "", sWhere)
          End If
          sSentencia = sSentencia & sWhereIni
          If (aTabla(nSecuencia) = "CoCpbCab" Or aTabla(nSecuencia) = "CoCpbDet") And Check.Value = 1 Then
            sSentencia = sSentencia & " and a.coddro " & IIf(checkincluir.Value = 1, "", "not ") & "in ('" & txtDato(0).Text & "')"
          End If
          With porstTmp
            If .State = adStateOpen Then .Close
            .ActiveConnection = pocnnMain
            .Source = sSentencia
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
            .Open
          End With
          If Not (porstTmp.BOF And porstTmp.EOF) Then
            ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
            lblProgreso(1).Caption = Choose(gsIdioma, "Exportando Archivo: ", "Exporting File: ") & Trim(chkImporProceso(nContador).Caption)
            nNumRegistros = porstTmp.RecordCount
            pgbProgreso(1).Max = nNumRegistros
            pgbProgreso(1).Value = pgbProgreso(1).Min
            nRegistro = 0
            sArchivo = dlbDirectorio.path & "\" & gsRUCEmp & aArchivo(nSecuencia) & ".txt"
            ' Elimino archivo de texto si existe
            If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
            If Dir$(sArchivo, vbNormal) = "" Then
              Open sArchivo For Output Access Write Lock Read Write As #nArchivo
              ' Grabo los nombres y tipos de los campos de la tabla seleccionada
              sMilinea = pfRegistro_Texto(porstTmp, aColumnas(nSecuencia), 0)
              Print #nArchivo, sMilinea
              sMilinea = pfRegistro_Texto(porstTmp, aColumnas(nSecuencia), 1)
              Print #nArchivo, sMilinea
              While Not porstTmp.EOF
                nRegistro = nRegistro + 1
                ' Diseño y grabro la linea en el archivo
                sMilinea = pfRegistro_Texto(porstTmp, aColumnas(nSecuencia), 2)
                Print #nArchivo, sMilinea
                pgbProgreso(1).Value = nRegistro
                DoEvents
                porstTmp.MoveNext
              Wend
              Close #nArchivo
            End If
          End If
          porstTmp.Close
        Next nSecuencia
      End If
    Next nContador
    Set porstTmp = Nothing

End Sub
'Function ftot_field_arr(xaTabla() As String) As Integer() 'As Long() '
'Function ftot_field_arr(ByRef xaTabla() As String) As Integer() 'As Long() '
'Function ftot_field_arr(xaTabla() As Variant) As Integer() 'As Long() '
Function ftot_field_arr(xaTabla() As Variant) As Variant 'As Long() '
'ini 2015-08-04 verificar estructura de la base de datos
        Dim yrst As ADODB.Recordset
        Dim yyTabla As String
        Dim xtot_field0 As Integer, xtot_field1 As Integer
        Dim xtot_field2 As Integer, xtot_field3 As Integer
        xtot_field0 = 0
        xtot_field1 = 0
        xtot_field2 = 0
        xtot_field3 = 0
        'Dim xaColumnas() As Integer
        Dim xaColumnas()
        Dim nSecuencia As Integer
        For nSecuencia = 0 To UBound(xaTabla, 1)
            'funcion para hallar los datos de la estructura
            'por ahora haya el total campos de CoCieMes
                yyTabla = xaTabla(nSecuencia)
                'Abro el recordset
                'Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, "CoCieMes"))
                Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, yyTabla))
                With yrst
                    If nSecuencia = 0 Then
                    xtot_field0 = .RecordCount
                    End If
                    If nSecuencia = 1 Then
                    xtot_field1 = .RecordCount
                    End If
                    If nSecuencia = 2 Then
                    xtot_field2 = .RecordCount
                    End If
                    If nSecuencia = 3 Then
                    xtot_field3 = .RecordCount
                    End If
                End With
                fRstClose yrst
        Next nSecuencia
'        If xtot_field2 = 0 Then
'            aColumnas = Array(xtot_field)
'        Else
'            aColumnas = Array(xtot_field, xtot_field2)
'        End If
        If xtot_field3 > 0 Then
            xaColumnas = Array(xtot_field0, xtot_field1, xtot_field2, xtot_field3)
        Else
            If xtot_field2 > 0 Then
              xaColumnas = Array(xtot_field0, xtot_field1, xtot_field2)
            Else
                If xtot_field1 > 0 Then
                  xaColumnas = Array(xtot_field0, xtot_field1)
                Else
                  xaColumnas = Array(xtot_field0)
                End If
            End If
        End If
'fin 2015-08-04 verificar estructura de la base de datos
   ftot_field_arr = xaColumnas

End Function


Private Sub ppBackup_Tablas()
    Dim sSentencia As String, sMilinea As String
    Dim nContador As Integer, nArchivo As Integer
    Dim nRegistro As Double, nNumRegistros As Double
    Dim sArchivo As String, nSecuencia As Integer
    Dim aTabla(), aArchivo(), aColumnas(), aWhere()
    Dim porstTmp As ADODB.Recordset

    ' Seteo el recordset temporal para la grabacion
    Set porstTmp = New ADODB.Recordset
    ' Importo las tablas de acuerdo a la selección
    For nContador = 0 To chkImporTabla.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporTabla(nContador).Value Then
        ' Obtengo el archivo de texto
        'aArchivo = Choose(nContador + 1, Array("dro"), Array("bco"), Array("pct"), Array("aux", "nat"), Array("tdc"), Array("cco"), Array("fef", "fjo"), Array("tpc", "tcm"), Array("cfg", "cfe"), Array("efi", "cef", "psp"), Array("rpt", "cms", "fic", "fid"), Array("asi", "asd"), Array("dpe"), Array("ot1", "ot2", "ot3", "ot4"))
        'aColumnas = Choose(nContador + 1, Array(24), Array(12), Array(34), Array(17, 11), Array(11), Array(11), Array(10, 11), Array(10, 10), Array(14, 33), Array(11, 24, 35), Array(13, 7, 10, 28), Array(10, 13), Array(9), Array(11, 10, 5, 8))
        ' Genero el archivo y abro el recordset temporal
        'aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("codpe"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
        'aWhere = Choose(nContador + 1, Array("AND pdoano='" & gsAnoAct & "'"), Array(""), Array("AND pdoano='" & gsAnoAct & "'"), Array("", ""), Array(""), Array("AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array(""), Array("", "", "", ""))
 '2014-08-20
'        aArchivo = Choose(nContador + 1, Array("dro"), Array("bco"), Array("pct"), Array("aux", "nat"), Array("tdc"), Array("cco"), Array("fef", "fjo"), Array("tpc", "tcm"), Array("cfg", "cfe"), Array("dpe"), Array("rpt", "cms", "fic", "fid"), Array("asi", "asd"), Array("efi", "cef", "psp"), Array("ot1", "ot2", "ot3", "ot4"))
'        aColumnas = Choose(nContador + 1, Array(24), Array(12), Array(35), Array(19, 11), Array(11), Array(11), Array(10, 11), Array(10, 10), Array(14, 36), Array(9), Array(13, 7, 10, 28), Array(10, 13), Array(11, 24, 33), Array(11, 10, 5, 8))
'        aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("codpe"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
'        aWhere = Choose(nContador + 1, Array("AND pdoano='" & gsAnoAct & "'"), Array(""), Array("AND pdoano='" & gsAnoAct & "'"), Array("", ""), Array(""), Array("AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array(""), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("", "", "", ""))
        
        
 '''ini 2015-08-04 verificar estructura de la base de datos
'''funcion para hallar los datos de la estructura
'''por ahora haya el total campos de CoCieMes
''    Dim yrst As ADODB.Recordset
''    'Abro el recordset
''    Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, "CoCieMes"))
''    Dim xtot_field As Integer
''    xtot_field = 0
''    With yrst
''            xtot_field = .RecordCount
''    End With
''    fRstClose yrst
'''fin 2015-08-04 verificar estructura de la base de datos
       
        
        aArchivo = Choose(nContador + 1, Array("dro"), Array("bco"), Array("pct"), Array("aux", "nat"), Array("tdc"), Array("cco"), Array("fef", "fjo"), Array("tpc", "tcm"), Array("cfg", "cfe"), Array("dpe"), Array("rpt", "cms", "fic", "fid"), Array("asi", "asd"), Array("efi", "cef", "psp"), Array("ot1", "ot2", "ot3", "ot4"))
        '2015-08-04 verificar estructura de la base de datos aColumnas = Choose(nContador + 1, Array(24), Array(12), Array(35), Array(19, 11), Array(11), Array(11), Array(10, 11), Array(10, 10), Array(15, 36), Array(9), Array(13, 7, 10, 28), Array(10, 13), Array(11, 24, 33), Array(11, 10, 5, 8))
        aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("codpe"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
        aWhere = Choose(nContador + 1, Array("AND pdoano='" & gsAnoAct & "'"), Array(""), Array("AND pdoano='" & gsAnoAct & "'"), Array("", ""), Array(""), Array("AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array(""), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'", "AND pdoano='" & gsAnoAct & "'"), Array("", "", "", ""))
        
'''ini 2015-08-04 verificar estructura de la base de datos
        aColumnas = ftot_field_arr(aTabla)
''        Dim yrst As ADODB.Recordset
''        Dim yyTabla As String
''        Dim xtot_field0 As Integer, xtot_field1 As Integer
''        Dim xtot_field2 As Integer, xtot_field3 As Integer
''        xtot_field0 = 0
''        xtot_field1 = 0
''        xtot_field2 = 0
''        xtot_field3 = 0
''        For nSecuencia = 0 To UBound(aArchivo, 1)
''            'funcion para hallar los datos de la estructura
''            'por ahora haya el total campos de CoCieMes
''                yyTabla = aTabla(nSecuencia)
''                'Abro el recordset
''                'Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, "CoCieMes"))
''                Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, yyTabla))
''                With yrst
''                    If nSecuencia = 0 Then
''                    xtot_field0 = .RecordCount
''                    End If
''                    If nSecuencia = 1 Then
''                    xtot_field1 = .RecordCount
''                    End If
''                    If nSecuencia = 2 Then
''                    xtot_field2 = .RecordCount
''                    End If
''                    If nSecuencia = 3 Then
''                    xtot_field3 = .RecordCount
''                    End If
''                End With
''                fRstClose yrst
''        Next nSecuencia
'''        If xtot_field2 = 0 Then
'''            aColumnas = Array(xtot_field)
'''        Else
'''            aColumnas = Array(xtot_field, xtot_field2)
'''        End If
''        If xtot_field3 > 0 Then
''            aColumnas = Array(xtot_field0, xtot_field1, xtot_field2, xtot_field3)
''        Else
''            If xtot_field2 > 0 Then
''              aColumnas = Array(xtot_field0, xtot_field1, xtot_field2)
''            Else
''                If xtot_field1 > 0 Then
''                  aColumnas = Array(xtot_field0, xtot_field1)
''                Else
''                  aColumnas = Array(xtot_field0)
''                End If
''            End If
''        End If
'''fin 2015-08-04 verificar estructura de la base de datos
        
        For nSecuencia = 0 To UBound(aArchivo, 1)
          ' Obtengo el archivo de texto libre
          nArchivo = FreeFile
          sSentencia = "SELECT * "
          sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " "
          sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
          sSentencia = sSentencia & aWhere(nSecuencia)
          With porstTmp
            If .State = adStateOpen Then .Close
            .ActiveConnection = pocnnMain
            .Source = sSentencia
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
            .Open
          End With
          If Not (porstTmp.BOF And porstTmp.EOF) Then
            ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
            lblProgreso(1).Caption = Choose(gsIdioma, "Exportando Archivo: ", "Exporting File: ") & Trim(chkImporTabla(nContador).Caption)
            nNumRegistros = porstTmp.RecordCount
            pgbProgreso(1).Max = nNumRegistros
            pgbProgreso(1).Value = pgbProgreso(1).Min
            nRegistro = 0
            sArchivo = dlbDirectorio.path & "\" & gsRUCEmp & aArchivo(nSecuencia) & ".txt"
            ' Elimino archivo de texto si existe
            If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
            If Dir$(sArchivo, vbNormal) = "" Then
              Open sArchivo For Output Access Write Lock Read Write As #nArchivo
              ' Grabo los nombres, tipos de los campos de la tabla seleccionada
              sMilinea = pfRegistro_Texto(porstTmp, aColumnas(nSecuencia), 0)
              Print #nArchivo, sMilinea
              sMilinea = pfRegistro_Texto(porstTmp, aColumnas(nSecuencia), 1)
              Print #nArchivo, sMilinea
              While Not porstTmp.EOF
                nRegistro = nRegistro + 1
                ' Diseño y grabro la linea en el archivo
                sMilinea = pfRegistro_Texto(porstTmp, aColumnas(nSecuencia), 2)
                Print #nArchivo, sMilinea
                pgbProgreso(1).Value = nRegistro
                DoEvents
                porstTmp.MoveNext
              Wend
              Close #nArchivo
            End If
          End If
          porstTmp.Close
        Next nSecuencia
      End If
    Next nContador
    Set porstTmp = Nothing

End Sub

Private Sub ppRegistro_Texto(ByVal sLinea As String, ByVal nColumnas As Integer, ByRef aRegistros, bTipo As Byte)
    
    Static nCampo As Integer
    Static nInicio As Integer, nLongitud As Integer
    ReDim Preserve aRegistros(3, nColumnas)
    nInicio = 1
    For nCampo = 1 To nColumnas
      nLongitud = Abs(InStr(nInicio, sLinea, "|") - nInicio)
      aRegistros(bTipo, nCampo) = Mid$(sLinea, nInicio, nLongitud)
      nInicio = nInicio + (nLongitud + 1)
    Next nCampo

End Sub

Private Sub ppRestore_Proceso()
    Dim sSentencia As String, sMilinea As String
    Dim nContador As Integer, nArchivo As Integer
    Dim nRegistro As Double, nNumRegistros As Double
    Dim sArchivo As String, nSecuencia As Integer
    Dim aTabla(), aArchivo(), aColumnas(), aRegistros(), aClave()
    Dim nIndex As Integer, sSenValues As String, sWhere As String

    ' Obtengo el archivo de texto libre
    nArchivo = FreeFile
    ' Importo las tablas de acuerdo a la selección
    For nContador = 0 To chkImporProceso.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporProceso(nContador).Value = vbChecked Then
        ' Obtengo el archivo de texto
        aArchivo = Choose(nContador + 1, Array("pdo", "pdc"), Array("svc"), Array("cpd", "cpc", "cpo"), Array("vtd", "vtc", "vto"), Array("hrd", "hrc", "hro"), Array("cbc", "cbd"), Array("cdc", "cdd", "cdr", "cdf"), Array("sct", "sax", "sco"))
        '2014-10-17 aumento cab pedido aColumnas = Choose(nContador + 1, Array(21, 14), Array(16), Array(88, 18, 16), Array(75, 19, 15), Array(45, 17, 15), Array(28, 26), Array(15, 37, 24, 16), Array(63, 64, 64))
        
        '2015-08-04 verificar estructura de la base de datos  aColumnas = Choose(nContador + 1, Array(22, 14), Array(16), Array(88, 18, 16), Array(75, 19, 15), Array(45, 17, 15), Array(28, 26), Array(15, 37, 24, 16), Array(63, 64, 64))
        ' Obtengo el nombre de la tabla
        aTabla = Choose(nContador + 1, Array("copdocpr", "copdocprcta"), Array("coconser"), Array("CoCprDoc", "CoCprDocCta", "CoCprDocCCo"), Array("CoVtaDoc", "CoVtaDocCta", "CoVtaDocCCo"), Array("CoHprDoc", "CoHprDocCta", "CoHprDocCCo"), Array("CoBanCab", "CoBanDet"), Array("CoCpbCab", "CoCpbDet", "CoCpbDetRP", "CoCpbDetFjo"), Array("coctaacu", "coauxacu", "coccoacu"))
        aClave = Choose(nContador + 1, Array("coddpe", "coddpe"), Array("codcon"), Array("CodAux", "CodAux", "CodAux"), Array("CodTDc", "CodTDc", "CodTDc"), Array("CodAux", "CodAux", "CodAux"), Array("CodDro", "CodDro"), Array("CodDro", "CodDro", "CodDro", "CodDro"), Array("codcta", "codcta", "codcta"))
        sWhere = Choose(nContador + 1, " AND b.mespvs=a.mespvs AND b.coddpe=a.coddpe AND b.pdocpr=a.pcocpr", " AND b.mespvs=a.mespvs AND b.codcon=a.codcon", " AND b.CodAux=a.CodAux AND b.CodTDc=a.CodTDc AND b.SerDoc=a.SerDoc AND b.NroDoc=a.NroDoc", " AND b.CodTDc=a.CodTDc AND b.SerDoc=a.SerDoc AND b.NroDoc=a.NroDoc", " AND b.CodAux=a.CodAux AND b.SerDoc=a.SerDoc AND b.NroDoc=a.NroDoc", "", "", "")
        
'''ini 2015-08-04 verificar estructura de la base de datos
        aColumnas = ftot_field_arr(aTabla)
'''fin 2015-08-04 verificar estructura de la base de datos
        
        ' Desactivo la opcion si no existe archivo
        chkImporProceso(nContador).Value = vbUnchecked
        For nSecuencia = 0 To UBound(aArchivo, 1)
        
    
        
          'Elimino y creo el archivo temporal de grabacion/restauración de información
          pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp" & aTabla(nSecuencia), "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, " & Len(aTabla(nSecuencia)) + 5 & ")='#tmp" & aTabla(nSecuencia) & "_') DROP TABLE #tmp" & aTabla(nSecuencia))
          sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmp" & aTabla(nSecuencia) & " ", "")
          sSentencia = sSentencia & "SELECT * "
          sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmp" & aTabla(nSecuencia) & " ", "")
          sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " "
          sSentencia = sSentencia & "WHERE " & aClave(nSecuencia) & "=Null"
          
          pocnnMain.Execute sSentencia
        
          sArchivo = dlbDirectorio.path & "\" & gsRUCEmp & aArchivo(nSecuencia) & ".txt"
          ' Verifico si existe el archivo de texto y activo la opción
          If Dir$(sArchivo, vbNormal) <> "" Then
            chkImporProceso(nContador).Value = vbChecked
            Open sArchivo For Input As #nArchivo
            nNumRegistros = gfRedond(LOF(nArchivo), 0)
            If nNumRegistros > 0 Then
              ' Barro todo el archivo de texto y grabo en la tabla temporal creada
              ReDim aRegistros(3, aColumnas(nSecuencia))
              pgbProgreso(1).Max = nNumRegistros
              pgbProgreso(1).Value = pgbProgreso(1).Min
              nRegistro = 0
              lblProgreso(1).Caption = Choose(gsIdioma, "Importando Archivo: ", "importing File: ") & Trim(chkImporProceso(nContador).Caption)
              ' Obtengo los nombres de los campos
              Line Input #nArchivo, sMilinea
              ppRegistro_Texto sMilinea, aColumnas(nSecuencia), aRegistros, 0
              nRegistro = 1
              
              ' Obtengo los tipos de los campos
              Line Input #nArchivo, sMilinea
              ppRegistro_Texto sMilinea, aColumnas(nSecuencia), aRegistros, 1
              nRegistro = 2
              
              ' Inserto los datos o detalle la tabla temporal
              Dim ax As String
              Do While Not EOF(nArchivo)
                Line Input #nArchivo, sMilinea
                nRegistro = nRegistro + 1
                ppRegistro_Texto sMilinea, aColumnas(nSecuencia), aRegistros, 2
                aRegistros(2, 1) = gsCodEmp
                sSentencia = "INSERT INTO " & ps_Prefijo & "tmp" & aTabla(nSecuencia) & "("
                sSenValues = "VALUES("
                For nIndex = 1 To aColumnas(nSecuencia)
                  cmdAceptar.Tag = IIf(nIndex = aColumnas(nSecuencia), ") ", ", ")
                  sSentencia = sSentencia & aRegistros(0, nIndex) & cmdAceptar.Tag
                  If aRegistros(1, nIndex) = "C" Then
                    sSenValues = sSenValues & IIf(aRegistros(2, nIndex) = "", "Null", "'" & aRegistros(2, nIndex) & "'")
                  ElseIf aRegistros(1, nIndex) = "F" Then
                    sSenValues = sSenValues & IIf(aRegistros(2, nIndex) = "", "Null", IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(aRegistros(2, nIndex), "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & ")")
                  ElseIf aRegistros(1, nIndex) = "T" Then
                    sSenValues = sSenValues & IIf(aRegistros(2, nIndex) = "", "Null", IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(aRegistros(2, nIndex), s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ")")
                  ElseIf aRegistros(1, nIndex) = "N" Then
                    sSenValues = sSenValues & Val(aRegistros(2, nIndex))
                  End If
                  sSenValues = sSenValues & cmdAceptar.Tag
                Next nIndex
                sSentencia = sSentencia & sSenValues
                ' Ejecuto la insercion del registro en la tabla
                pocnnMain.Execute sSentencia
                
                ' Actualizo la barra de progreso
                pgbProgreso(1).Value = IIf((Loc(nArchivo) * 128) > nNumRegistros, nNumRegistros, (Loc(nArchivo) * 128))
                DoEvents
              Loop
            End If
            Close #nArchivo
          End If
        Next nSecuencia
      End If
    Next nContador

End Sub

Private Sub ppRestore_Tablas()
    
    Dim sSentencia As String, sMilinea As String
    Dim nContador As Integer, nArchivo As Integer
    Dim nRegistro As Double, nNumRegistros As Double
    Dim sArchivo As String, nSecuencia As Integer
    Dim aTabla(), aArchivo(), aColumnas(), aRegistros()
    Dim nIndex As Integer, sSenValues As String, aClave()

    ' Obtengo el archivo de texto libre
    nArchivo = FreeFile
    ' Importo las tablas de acuerdo a la selección
    For nContador = 0 To chkImporTabla.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporTabla(nContador).Value Then
        ' Obtengo el archivo de texto
        'aArchivo = Choose(nContador + 1, Array("dro"), Array("bco"), Array("pct"), Array("aux", "nat"), Array("tdc"), Array("cco"), Array("fef", "fjo"), Array("tpc", "tcm"), Array("cfg", "cfe"), Array("efi", "cef", "psp"), Array("rpt", "cms", "fic", "fid"), Array("asi", "asd"), Array("dpe"), Array("ot1", "ot2", "ot3", "ot4"))
        'aColumnas = Choose(nContador + 1, Array(24), Array(12), Array(34), Array(17, 11), Array(11), Array(11), Array(10, 11), Array(10, 10), Array(14, 33), Array(11, 24, 35), Array(13, 7, 10, 28), Array(10, 10), Array(9), Array(11, 10, 5, 8))
        ' Obtengo el nombre de la tabla
        'aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("codpe"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
        'aClave = Choose(nContador + 1, Array("CodDro"), Array("codbco"), Array("CodCta"), Array("CodAux", "CodAux"), Array("CodTDc"), Array("CodCCo"), Array("CodEfe", "CodFjo"), Array("FehTCb", "MesPvs"), Array("pdoano", "pdoano"), Array("CodEfi", "CodEfi", "CodCta"), Array("NumOrd", "MesCie", "codfil", "codfil"), Array("codasi", "codasi"), Array("coddpe"), Array("codbco", "codmed", "codemp", "codlib"))
        
          '2014-08-20 aColumnas = Choose(nContador + 1, Array(24), Array(12), Array(35), Array(19, 11), Array(11), Array(11), Array(10, 11), Array(10, 10), Array(14, 36), Array(9), Array(13, 7, 10, 28), Array(10, 13), Array(11, 24, 33), Array(11, 10, 5, 8))
      
        aArchivo = Choose(nContador + 1, Array("dro"), Array("bco"), Array("pct"), Array("aux", "nat"), Array("tdc"), Array("cco"), Array("fef", "fjo"), Array("tpc", "tcm"), Array("cfg", "cfe"), Array("dpe"), Array("rpt", "cms", "fic", "fid"), Array("asi", "asd"), Array("efi", "cef", "psp"), Array("ot1", "ot2", "ot3", "ot4"))
        '2015-08-04 verificar estructura de la base de datos aColumnas = Choose(nContador + 1, Array(24), Array(12), Array(35), Array(19, 11), Array(11), Array(11), Array(10, 11), Array(10, 10), Array(15, 36), Array(9), Array(13, 7, 10, 28), Array(10, 13), Array(11, 24, 33), Array(11, 10, 5, 8))
        aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("codpe"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
        aClave = Choose(nContador + 1, Array("CodDro"), Array("codbco"), Array("CodCta"), Array("CodAux", "CodAux"), Array("CodTDc"), Array("CodCCo"), Array("CodEfe", "CodFjo"), Array("FehTCb", "MesPvs"), Array("pdoano", "pdoano"), Array("coddpe"), Array("NumOrd", "MesCie", "codfil", "codfil"), Array("codasi", "codasi"), Array("CodEfi", "CodEfi", "CodCta"), Array("codbco", "codmed", "codemp", "codlib"))
'''ini 2015-08-04 verificar estructura de la base de datos
        aColumnas = ftot_field_arr(aTabla)
'''fin 2015-08-04 verificar estructura de la base de datos

        
        ' Desactivo la opcion si no existe archivo
        chkImporTabla(nContador).Value = vbUnchecked
        For nSecuencia = 0 To UBound(aArchivo, 1)
          ' Elimino y creo el archivo temporal de grabacion/restauración de información
          pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp" & aTabla(nSecuencia), "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, " & Len(aTabla(nSecuencia)) + 5 & ")='#tmp" & aTabla(nSecuencia) & "_') DROP TABLE #tmp" & aTabla(nSecuencia))
          sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmp" & aTabla(nSecuencia) & " ", "")
          sSentencia = sSentencia & "SELECT * "
          sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmp" & aTabla(nSecuencia) & " ", "")
          sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " "
          sSentencia = sSentencia & "WHERE " & aClave(nSecuencia) & "=Null"
          pocnnMain.Execute sSentencia
          
          sArchivo = dlbDirectorio.path & "\" & gsRUCEmp & aArchivo(nSecuencia) & ".txt"
          ' Verifico si existe el archivo de texto y activo la opción
          If Dir$(sArchivo, vbNormal) <> "" Then
            chkImporTabla(nContador).Value = vbChecked
            Open sArchivo For Input As #nArchivo
            nNumRegistros = gfRedond(LOF(nArchivo), 0)
            If nNumRegistros > 0 Then
              ' Barro todo el archivo de texto y grabo en la tabla temporal creada
              ReDim aRegistros(3, aColumnas(nSecuencia))
              pgbProgreso(1).Max = nNumRegistros
              pgbProgreso(1).Value = pgbProgreso(1).Min
              nRegistro = 0
              lblProgreso(1).Caption = Choose(gsIdioma, "Importando Archivo: ", "importing File: ") & Trim(chkImporTabla(nContador).Caption)
              ' Obtengo los nombres de los campos
              Line Input #nArchivo, sMilinea
              ppRegistro_Texto sMilinea, aColumnas(nSecuencia), aRegistros, 0
              nRegistro = 1
              
              ' Obtengo los tipos de los campos
              Line Input #nArchivo, sMilinea
              ppRegistro_Texto sMilinea, aColumnas(nSecuencia), aRegistros, 1
              nRegistro = 2
              
              ' Inserto los datos a la tabla temporal
              Do While Not EOF(nArchivo)
                Line Input #nArchivo, sMilinea
                nRegistro = nRegistro + 1
                ppRegistro_Texto sMilinea, aColumnas(nSecuencia), aRegistros, 2
                aRegistros(2, 1) = gsCodEmp
                sSentencia = "INSERT INTO " & ps_Prefijo & "tmp" & aTabla(nSecuencia) & "("
                sSenValues = "VALUES("
                For nIndex = 1 To aColumnas(nSecuencia)
                  cmdAceptar.Tag = IIf(nIndex = aColumnas(nSecuencia), ") ", ", ")
                  sSentencia = sSentencia & aRegistros(0, nIndex) & cmdAceptar.Tag
                  If aRegistros(1, nIndex) = "C" Then
                    sSenValues = sSenValues & IIf(aRegistros(2, nIndex) = "", "Null", "'" & aRegistros(2, nIndex) & "'")
                  ElseIf aRegistros(1, nIndex) = "F" Then
                    sSenValues = sSenValues & IIf(aRegistros(2, nIndex) = "", "Null", IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(aRegistros(2, nIndex), "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & ")")
                  ElseIf aRegistros(1, nIndex) = "T" Then
                    sSenValues = sSenValues & IIf(aRegistros(2, nIndex) = "", "Null", IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(aRegistros(2, nIndex), s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ")")
                  ElseIf aRegistros(1, nIndex) = "N" Then
                    sSenValues = sSenValues & CDec(aRegistros(2, nIndex))
                  End If
                  sSenValues = sSenValues & cmdAceptar.Tag
                Next nIndex
                sSentencia = sSentencia & sSenValues
                ' Ejecuto la insercion del registro en la tabla
                pocnnMain.Execute sSentencia
                
                ' Actualizo la barra de progreso
                pgbProgreso(1).Value = IIf((Loc(nArchivo) * 128) > nNumRegistros, nNumRegistros, (Loc(nArchivo) * 128))
                DoEvents
              Loop
            End If
            Close #nArchivo
          End If
        Next nSecuencia
      End If
    Next nContador

End Sub

Private Sub ppValida_Proceso()
    
    Dim sSentencia As String, sMilinea As String
    Dim nContador As Integer, nNumRegistros As Double
    Dim nSecuencia As Integer, aTabla()
    Dim aJoin(), aWhere(), aOrden()
    Dim sWhere As String, sCampo As String
    Dim porstTmp As ADODB.Recordset
    Dim aCampos(2, 2) As String, nIndex As Integer
    
    ' Seteo el recordset temporal para la grabacion
    Set porstTmp = New ADODB.Recordset

    ' Importo las tablas de acuerdo a la selección
    For nContador = 0 To chkImporProceso.Count - 1
    ' Verifico que se haya seleccionado
      If chkImporProceso(nContador).Value Then
      
        'Obtengo el nombre de la tabla
        aTabla = Choose(nContador + 1, Array("copdocpr", "copdocprcta"), Array("coconser"), Array("CoCprDoc", "CoCprDocCta", "CoCprDocCCo"), Array("CoVtaDoc", "CoVtaDocCta", "CoVtaDocCCo"), Array("CoHprDoc", "CoHprDocCta", "CoHprDocCCo"), Array("CoBanCab", "CoBanDet"), Array("CoCpbCab", "CoCpbDet", "CoCpbDetRP", "CoCpbDetFjo"), Array("CoCtaAcu", "CoAuxAcu", "CoCCoAcu"))
        'aTabla = Choose(nContador + 1, Array("copdocpr"), Array("CoCprDoc", "CoCprDocCta", "CoCprDocCCo"), Array("CoVtaDoc", "CoVtaDocCta", "CoVtaDocCCo"), Array("CoHprDoc", "CoHprDocCta", "CoHprDocCCo"), Array("CoBanCab", "CoBanDet"), Array("CoCpbCab", "CoCpbDet", "CoCpbDetRP", "CoCpbDetFjo"), Array("CoCtaAcu", "CoAuxAcu", "CoCCoAcu"))
        aJoin = Choose(nContador + 1, Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.coddpe=b.coddpe AND a.pdocpr=b.pdocpr", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.coddpe=b.coddpe AND a.pdocpr=b.pdocpr and a.codcta=b.codcta and a.codcco=b.codcco"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.codcon=b.codcon"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc AND a.TpoCnc=b.TpoCnc AND a.Orden=b.Orden AND a.CodCta=b.CodCta", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc AND a.TpoCnc=b.TpoCnc AND a.Orden=b.Orden AND a.CodCta=b.CodCta AND a.CodCCo=b.CodCCo"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc AND a.TpoCnc=b.TpoCnc AND a.Orden=b.Orden AND a.CodCta=b.CodCta", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc AND a.TpoCnc=b.TpoCnc AND a.Orden=b.Orden AND a.CodCta=b.CodCta AND a.CodCco=b.CodCco"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc AND a.TpoCnc=b.TpoCnc AND a.Orden=b.Orden AND a.CodCta=b.CodCta", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodAux=b.CodAux AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc AND a.TpoCnc=b.TpoCnc AND a.Orden=b.Orden AND a.CodCta=b.CodCta AND a.CodCco=b.CodCco"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroBan=b.NroBan", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroBan=b.NroBan AND a.NroItem=b.NroItem"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte ", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte"), _
                                      Array("a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta ", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta AND a.CodAux=b.CodAux ", "a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta AND a.CodCCo=b.CodCCo"))
        'aWhere = Choose(nContador + 1, Array("a.mespvs, a.coddpe, a.pdocpr"), Array("a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc", "a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta", "a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta, a.CodCCo"), Array("a.CodTDc, a.SerDoc, a.NroDoc", "a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta", "a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta, a.CodCco"), Array("a.CodAux, a.SerDoc, a.NroDoc", "a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta", "a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta, a.CodCco"), Array("a.MesPvs, a.CodDro, a.NroBan", "a.MesPvs, a.CodDro, a.NroBan, RTrim(a.NroItem)"), Array("a.MesPvs, a.CodDro, a.NroCpb", "a.MesPvs, a.CodDro, a.NroCpb, RTrim(a.NroIte)", "a.MesPvs, a.CodDro, a.NroCpb, RTrim(a.NroIte)", "a.MesPvs, a.CodDro, a.NroCpb, RTrim(a.NroIte)"), Array("a.CodCta", "a.CodCta, a.CodAux", "a.CodCta, a.CodCCo"))
        aWhere = Choose(nContador + 1, Array("a.mespvs, a.coddpe, a.pdocpr", "a.mespvs, a.coddpe, a.pdocpr, a.codcta, a.codcco"), _
                                       Array("a.mespvs, a.codcon"), _
                                       Array("a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc", "a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta", "a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta, a.CodCCo"), _
                                       Array("a.CodTDc, a.SerDoc, a.NroDoc", "a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta", "a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta, a.CodCco"), Array("a.CodAux, a.SerDoc, a.NroDoc", "a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta", "a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, RTrim(a.Orden), a.CodCta, a.CodCco"), Array("a.MesPvs, a.CodDro, a.NroBan", "a.MesPvs, a.CodDro, a.NroBan, RTrim(a.NroItem)"), Array("a.MesPvs, a.CodDro, a.NroCpb", "a.MesPvs, a.CodDro, a.NroCpb, RTrim(a.NroIte)", "a.MesPvs, a.CodDro, a.NroCpb, RTrim(a.NroIte)", "a.MesPvs, a.CodDro, a.NroCpb, RTrim(a.NroIte)"), Array("a.CodCta", "a.CodCta, a.CodAux", "a.CodCta, a.CodCCo"))
        
        aOrden = Choose(nContador + 1, Array("a.mespvs, a.coddpe, a.pdocpr", "a.mespvs, a.coddpe, a.pdocpr"), Array("a.mespvs, a.codcon"), Array("a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc", "a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCta", "a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCta, a.CodCCo"), Array("a.CodTDc, a.SerDoc, a.NroDoc", "a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCta", "a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCta, a.CodCco"), Array("a.CodAux, a.SerDoc, a.NroDoc", "a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCta", "a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCta, a.CodCco"), Array("a.MesPvs, a.CodDro, a.NroBan", "a.MesPvs, a.CodDro, a.NroBan, a.NroItem"), Array("a.MesPvs, a.CodDro, a.NroCpb", "a.MesPvs, a.CodDro, a.NroCpb, a.NroIte", "a.MesPvs, a.CodDro, a.NroCpb, a.NroIte", "a.MesPvs, a.CodDro, a.NroCpb, a.NroIte, a.NroOrd"), Array("a.CodCta", "a.CodCta, a.CodAux", "a.CodCta, a.CodCCo"))
        sWhere = Choose(nContador + 1, "a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.mespvs=c.mespvs AND a.coddpe=c.coddpe AND a.pdocpr=c.pdocpr ", "a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.mespvs=c.mespvs AND a.codcon=c.codcon ", "a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodAux=c.CodAux AND a.CodTDc=c.CodTDc AND a.SerDoc=c.SerDoc AND a.NroDoc=c.NroDoc ", "a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodTDc=c.CodTDc AND a.SerDoc=c.SerDoc AND a.NroDoc=c.NroDoc ", "a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodAux=c.CodAux AND a.SerDoc=c.SerDoc AND a.NroDoc=c.NroDoc ", "", "", "")
                
          
        For nSecuencia = 0 To UBound(aTabla, 1)
        
          pgbProgreso(1).Max = IIf(UBound(aTabla, 1) > 1, UBound(aTabla, 1), 1)
          pgbProgreso(1).Value = pgbProgreso(1).Min
          lblProgreso(1).Caption = Choose(gsIdioma, "Importando Archivo: ", "importing File: ") & Trim(chkImporProceso(nContador).Caption)
          
          ' Inserto la información no existente a la tabla
          sMilinea = IIf(nSecuencia = 0 Or (nContador = 5 Or nContador = 6), "a", "c")
          
           
          aWhere(nSecuencia) = Replace(aWhere(nSecuencia), ", ", IIf(ps_Plataforma = pSrvMySql, ", ", "+"))
          
          If nContador <> 7 Or (nContador = 6 And chkProceso.Value = vbChecked) Then
            sSentencia = "INSERT INTO " & aTabla(nSecuencia) & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & aTabla(nSecuencia) & " a "
            sSentencia = sSentencia & IIf(nSecuencia = 0 Or (nContador = 5 Or nContador = 6), "", ", " & ps_Prefijo & "tmp" & aTabla(0) & " c ")
            sSentencia = sSentencia & "WHERE " & sMilinea & ".codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND " & sMilinea & ".pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(", "ISNULL((") & aWhere(nSecuencia) & "), '')<>'' "
            sSentencia = sSentencia & IIf(nSecuencia = 0 Or (nContador = 5 Or nContador = 6), "", "AND " & sWhere)
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT b.* FROM " & aTabla(nSecuencia) & " b "
            sSentencia = sSentencia & "WHERE " & aJoin(nSecuencia) & ") "
            If chkProceso.Value = vbUnchecked Then
              sSentencia = sSentencia & "AND " & sMilinea & ".MesPvs>='" & gfCeros(cmbPeriodo(0).ListIndex, 2, 0, "0") & "' "
              sSentencia = sSentencia & "AND " & sMilinea & ".MesPvs<='" & gfCeros(cmbPeriodo(1).ListIndex, 2, 0, "0") & "' "
            End If
            sSentencia = sSentencia & "ORDER BY " & aOrden(nSecuencia)
            pocnnMain.Execute sSentencia, nNumRegistros
         Else
            sSentencia = "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & aTabla(nSecuencia) & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "ORDER BY " & aOrden(nSecuencia)
            ' Actualizo o Inserto de acuerdo al parametro
            If chkProceso.Value = vbUnchecked Then
              ' Abro el recordset para actualizar los saldos
              With porstTmp
                If .State = adStateOpen Then .Close
                .ActiveConnection = pocnnMain
                .Source = sSentencia
                .CursorType = adOpenDynamic
                .LockType = adLockReadOnly
                .Open
              End With
              If Not (porstTmp.BOF And porstTmp.EOF) Then
                While Not porstTmp.EOF
                  ' Obtengo los campos a actualizar si no es general
                  aCampos(1, 1) = "": aCampos(1, 2) = ""
                  aCampos(2, 1) = "": aCampos(2, 2) = ""
                  For nIndex = cmbPeriodo(0).ListIndex To cmbPeriodo(1).ListIndex
                    sCampo = "AcuD" & Format(nIndex, "00") & "_MN"
                    aCampos(1, 1) = aCampos(1, 1) & sCampo & "=" & CDec(porstTmp(sCampo)) & ", "
                    sCampo = "AcuH" & Format(nIndex, "00") & "_MN"
                    aCampos(1, 2) = aCampos(1, 2) & sCampo & "=" & CDec(porstTmp(sCampo)) & ", "
                    sCampo = "AcuD" & Format(nIndex, "00") & "_ME"
                    aCampos(2, 1) = aCampos(2, 1) & sCampo & "=" & CDec(porstTmp(sCampo)) & ", "
                    sCampo = "AcuH" & Format(nIndex, "00") & "_ME"
                    aCampos(2, 2) = aCampos(2, 2) & sCampo & "=" & CDec(porstTmp(sCampo)) & IIf(nIndex = cmbPeriodo(1).ListIndex, "", ", ")
                  Next nIndex
                  ' Actualizo los registros con las columnas deseadas
                  sCampo = Choose(nSecuencia, "codcta", "codaux", "codcco")
                  sSentencia = "UPDATE " & aTabla(nSecuencia) & " SET "
                  sSentencia = sSentencia & aCampos(1, 1) & aCampos(1, 2)
                  sSentencia = sSentencia & aCampos(2, 1) & aCampos(2, 2) & " "
                  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
                  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
                  sSentencia = sSentencia & "AND codcta='" & porstTmp!CodCta & "' "
                  sSentencia = sSentencia & "AND " & sCampo & "='" & porstTmp(sCampo) & "'"
                  pocnnMain.Execute sSentencia, nNumRegistros
                  DoEvents
                  porstTmp.MoveNext
                Wend
              End If
              porstTmp.Close
            Else
              ' Inserto los registros restore general
              sSentencia = "INSERT INTO " & aTabla(nSecuencia) & " " & sSentencia
              pocnnMain.Execute sSentencia, nNumRegistros
            End If
          End If
          ' Actualizo la barra de progreso
          pgbProgreso(1).Value = IIf(nSecuencia > 1, nSecuencia, 1)
          DoEvents
        Next nSecuencia
      End If
    Next nContador
    Set porstTmp = Nothing

End Sub

Private Sub ppValida_Tablas()
    Dim sSentencia As String, sMilinea As String
    Dim nContador As Integer, nNumRegistros As Double
    Dim nSecuencia As Integer, aTabla()
    Dim aJoin(), aWhere(), aOrden(), aPeriodo()

    ' Importo las tablas de acuerdo a la selección
    For nContador = 0 To chkImporTabla.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporTabla(nContador).Value Then
        ' Obtengo el nombre de la tabla
        'aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("codpe"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
        'aJoin = Choose(nContador + 1, Array("b.CodDro=a.CodDro"), Array("b.codbco=a.codbco"), Array("b.CodCta=a.CodCta"), Array("b.CodAux=a.CodAux", "b.CodAux=a.CodAux"), Array("b.CodTdc=a.CodTdc"), Array("b.CodCCo=a.CodCCo"), Array("b.CodEfe=a.CodEfe", "b.CodFjo=a.CodFjo"), Array("b.FehTcb=a.FehTcb", "b.MesPvs=a.MesPvs"), Array("", ""), Array("b.CodEfi=a.CodEfi", "b.CodEfi=a.CodEfi AND b.NroLin=a.NroLin", "b.CodCta=a.CodCta"), Array("b.TipoFmt=a.TipoFmt AND b.NumOrd=a.NumOrd AND b.CodCCo=a.CodCCo", "b.MesCie=a.MesCie", "b.codfil=a.codfil", "b.codfil=a.codfil AND b.nrolin=a.nrolin"), Array("a.codemp=b.codemp AND a.CodAsi=b.CodAsi", "a.codemp=b.codemp AND a.CodAsi=b.CodAsi AND a.TpoCnc=b.TpoCnc AND a.codcta_mn=b.codcta_mn AND a.orden=b.orden"), Array("b.coddpe=a.coddpe"), Array("a.codaux=b.codaux and a.codbco=b.codbco and a.tpocta=b.tpocta and a.tpomon=b.tpomon ", "a.codmed=b.codmed", "a.proceso=b.proceso and a.valor=b.valor", "a.codlib=b.codlib"))
        'aWhere = Choose(nContador + 1, Array("a.CodDro"), Array("a.codbco"), Array("a.CodCta"), Array("a.CodAux", "a.CodAux"), Array("a.CodTDc"), Array("a.CodCCo"), Array("a.CodEfe", "a.CodFjo"), Array("a.FehTCb", "a.MesPvs"), Array("a.MesAtu", "a.MesAtu"), Array("a.CodEfi", "a.CodEfi, a.NroLin", "a.CodCta"), Array("a.NumOrd, a.CodCCo", "a.MesCie", "a.codfil", "a.codfil, a.nrolin"), Array("a.CodAsi", "a.CodAsi, a.TpoCnc, a.Codcta_mn, RTrim(a.orden)"), Array("a.coddpe"), Array("a.codbco", "a.codmed", "a.proceso", "a.codlib"))
        'aPeriodo = Choose(nContador + 1, Array("S"), Array("N"), Array("S"), Array("N", "N"), Array("N"), Array("S"), Array("S", "S"), Array("N", "S"), Array("S", "S"), Array("S", "S", "S"), Array("S", "S", "S", "S"), Array("S", "S"), Array("N"), Array("N", "N", "N", "N"))
        'aOrden = Choose(nContador + 1, Array("a.CodDro"), Array("a.codbco"), Array("a.CodCta"), Array("a.CodAux", "a.CodAux"), Array("a.CodTDc"), Array("a.CodCCo"), Array("a.CodEfe", "a.CodFjo"), Array("a.FehTCb", "a.MesPvs"), Array("a.pdoano, a.MesAtu", "a.pdoano, a.MesAtu"), Array("a.CodEfi", "a.CodEfi, a.NroLin", "a.CodCta"), Array("a.NumOrd, a.CodCCo", "a.MesCie", "a.codfil", "a.codfil, a.nrolin"), Array("a.CodAsi", "a.CodAsi, a.TpoCnc, a.Codcta_mn, a.orden"), Array("a.coddpe"), Array("a.codaux,a.codbco,a.tpocta,a.tpomon", "a.codmed", "a.proceso", "a.codlib"))
        
        aTabla = Choose(nContador + 1, Array("CoDro"), Array("cobco"), Array("CoCta"), Array("TgAux", "TgAuxNat"), Array("TgTDc"), Array("CoCCo"), Array("CoEfe", "CoFjo"), Array("TgTcb", "CoTcbMes"), Array("TgCfg", "CoCfg"), Array("codpe"), Array("CoCCoCfg", "CoCieMes", "cofil", "cofildet"), Array("coasitipo", "coasidet"), Array("CoEFi", "CoEFiLin", "CoPsp"), Array("coctaban", "bnmediopago", "rangoimpresion", "colib"))
        aJoin = Choose(nContador + 1, Array("b.CodDro=a.CodDro"), Array("b.codbco=a.codbco"), Array("b.CodCta=a.CodCta"), Array("b.CodAux=a.CodAux", "b.CodAux=a.CodAux"), Array("b.CodTdc=a.CodTdc"), Array("b.CodCCo=a.CodCCo"), Array("b.CodEfe=a.CodEfe", "b.CodFjo=a.CodFjo"), Array("b.FehTcb=a.FehTcb", "b.MesPvs=a.MesPvs"), Array("", ""), Array("b.coddpe=a.coddpe"), Array("b.TipoFmt=a.TipoFmt AND b.NumOrd=a.NumOrd AND b.CodCCo=a.CodCCo", "b.MesCie=a.MesCie", "b.codfil=a.codfil", "b.codfil=a.codfil AND b.nrolin=a.nrolin"), Array("a.codemp=b.codemp AND a.CodAsi=b.CodAsi", "a.codemp=b.codemp AND a.CodAsi=b.CodAsi AND a.TpoCnc=b.TpoCnc AND a.codcta_mn=b.codcta_mn AND a.orden=b.orden"), Array("b.CodEfi=a.CodEfi", "b.CodEfi=a.CodEfi AND b.NroLin=a.NroLin", "b.CodCta=a.CodCta"), Array("a.codaux=b.codaux and a.codbco=b.codbco and a.tpocta=b.tpocta and a.tpomon=b.tpomon ", "a.codmed=b.codmed", "a.proceso=b.proceso and a.valor=b.valor", "a.codlib=b.codlib"))
        aWhere = Choose(nContador + 1, Array("a.CodDro"), Array("a.codbco"), Array("a.CodCta"), Array("a.CodAux", "a.CodAux"), Array("a.CodTDc"), Array("a.CodCCo"), Array("a.CodEfe", "a.CodFjo"), Array("a.FehTCb", "a.MesPvs"), Array("a.MesAtu", "a.MesAtu"), Array("a.coddpe"), Array("a.NumOrd, a.CodCCo", "a.MesCie", "a.codfil", "a.codfil, a.nrolin"), Array("a.CodAsi", "a.CodAsi, a.TpoCnc, a.Codcta_mn, RTrim(a.orden)"), Array("a.CodEfi", "a.CodEfi, a.NroLin", "a.CodCta"), Array("a.codbco", "a.codmed", "a.proceso", "a.codlib"))
        aPeriodo = Choose(nContador + 1, Array("S"), Array("N"), Array("S"), Array("N", "N"), Array("N"), Array("S"), Array("S", "S"), Array("N", "S"), Array("S", "S"), Array("N"), Array("S", "S", "S", "S"), Array("S", "S"), Array("S", "S", "S"), Array("N", "N", "N", "N"))
        aOrden = Choose(nContador + 1, Array("a.CodDro"), Array("a.codbco"), Array("a.CodCta"), Array("a.CodAux", "a.CodAux"), Array("a.CodTDc"), Array("a.CodCCo"), Array("a.CodEfe", "a.CodFjo"), Array("a.FehTCb", "a.MesPvs"), Array("a.pdoano, a.MesAtu", "a.pdoano, a.MesAtu"), Array("a.coddpe"), Array("a.NumOrd, a.CodCCo", "a.MesCie", "a.codfil", "a.codfil, a.nrolin"), Array("a.CodAsi", "a.CodAsi, a.TpoCnc, a.Codcta_mn, a.orden"), Array("a.CodEfi", "a.CodEfi, a.NroLin", "a.CodCta"), Array("a.codaux,a.codbco,a.tpocta,a.tpomon", "a.codmed", "a.proceso", "a.codlib"))
        
        For nSecuencia = 0 To UBound(aTabla, 1)
          pgbProgreso(1).Max = IIf(UBound(aTabla, 1) > 1, UBound(aTabla, 1), 1)
          pgbProgreso(1).Value = pgbProgreso(1).Min
          lblProgreso(1).Caption = Choose(gsIdioma, "Importando Archivo: ", "importing File: ") & Trim(chkImporTabla(nContador).Caption)
          aWhere(nSecuencia) = Replace(aWhere(nSecuencia), ", ", IIf(ps_Plataforma = pSrvMySql, ", ", "+"))
          
          ' Inserto la información no existente a la tabla
          sSentencia = "INSERT INTO " & aTabla(nSecuencia) & " "
          sSentencia = sSentencia & "SELECT DISTINCT a.* "
          sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & aTabla(nSecuencia) & " a "
          sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
          sSentencia = sSentencia & IIf(aPeriodo(nSecuencia) = "S", "AND a.pdoano='" & gsAnoAct & "' ", "")
          sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(", "ISNULL((") & aWhere(nSecuencia) & "), '')<>'' "
          sSentencia = sSentencia & "AND NOT EXISTS (SELECT b.* FROM " & aTabla(nSecuencia) & " b "
          sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
          sSentencia = sSentencia & IIf(aPeriodo(nSecuencia) = "S", "AND b.pdoano=a.pdoano ", "")
          sSentencia = sSentencia & IIf(aJoin(nSecuencia) = "", "", "AND " & aJoin(nSecuencia)) & ") "
          sSentencia = sSentencia & "ORDER BY " & aOrden(nSecuencia)
          pocnnMain.Execute sSentencia, nNumRegistros
          ' Actualizo la barra de progreso
          pgbProgreso(1).Value = IIf(nSecuencia > 1, nSecuencia, 1)
          DoEvents
        Next nSecuencia
      End If
    Next nContador

End Sub

Private Sub drvUnidad_Change()

dlbDirectorio.path = drvUnidad.Drive
dlbDirectorio.Refresh

End Sub

Private Sub Form_Activate()
  cmdSalir.SetFocus
End Sub

Private Sub Form_Load()

Dim i As Integer

Check.Enabled = True
cmdDatoAyud(0).Enabled = False

Set pocnnMain = New ADODB.Connection
Set porstCodro = New ADODB.Recordset
   
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With

drvUnidad.Drive = gsRutSis
dlbDirectorio.path = gsRutSis

For i = 0 To 13
  If gsIdioma = NvlUsr_Sup Then
    cmbPeriodo(0).AddItem Choose(i + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre") & " " & gsAnoAct
    cmbPeriodo(1).AddItem Choose(i + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre") & " " & gsAnoAct
  Else
    cmbPeriodo(0).AddItem Choose(i + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing") & " " & gsAnoAct
    cmbPeriodo(1).AddItem Choose(i + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing") & " " & gsAnoAct
  End If
Next i
cmbPeriodo(0).ListIndex = Val(gsMesAct)
cmbPeriodo(1).ListIndex = Val(gsMesAct)
optProceso(0).Value = vbChecked


  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(4, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Directorio :", "Rango de periodos :", "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Directory :", "Range of periods :", "Home :", "End :")
  Next nElemento
  tabProceso.TabCaption(0) = Choose(gsIdioma, "Configuración de Parametros", "Configuration of Parameters")
  frmTablas.Caption = Choose(gsIdioma, " Tablas ", " Tables ")
  For nElemento = 0 To chkImporTabla.Count - 1
    If gsIdioma = NvlUsr_Sup Then
      chkImporTabla(nElemento).Caption = Choose(nElemento + 1, "&Diario", "&Bancos", "&Plan de Cuentas", "&Auxiliares", "&Tipo de Documentos", "&Centro de Costo", "&Flujo de Caja/Efectivo", "Tipo de Ca&mbio", "Ta&bla de Configuración", "&EE Financieros Presupuesto", "Configuración &Reportes", "A&siento Tipo", "&Proyecto", "&Otros")
    Else
      chkImporTabla(nElemento).Caption = Choose(nElemento + 1, "&Journal", "&Banks", "&Plan of Account", "&Auxiliaries", "&Type of Documents", "&Cost Center", "Cash/Money &Flow", "&Rate of Exchange", "Ta&ble of Configutaion", "Financial &Statement, Budget", "Configuraton &Reports", "&Standar Recorded", "&Project", "&Other")
    End If
  Next nElemento
  frmTransacciones.Caption = Choose(gsIdioma, " Transacciones ", " Transactions ")
  For nElemento = 0 To chkImporProceso.Count - 1
    If gsIdioma = NvlUsr_Sup Then
      chkImporProceso(nElemento).Caption = Choose(nElemento + 1, "&Pedidos de Compras", "Se&rvicios de Ventas", "Registro de &Compras", "Registro de &Ventas", "Registro de &Honorarios", "Caja &Bancos", "Comprobantes de &Diario", "&Saldos de Cuentas")
    Else
      chkImporProceso(nElemento).Caption = Choose(nElemento + 1, "Orders of &Purchase", "Sales Se&rvices", "P&urchase Register", "&Sales Register", "&Feed Register", "Cash and &Banks", "&Journal Vouchers", "&Balance of Accounts")
    End If
  Next nElemento
  frmUbicacion.Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  frmProceso.Caption = Choose(gsIdioma, " Tipo de Proceso ", " Type of Process ")
  optProceso(0).Caption = Choose(gsIdioma, "&Backup de Información", "&Backup of Information")
  optProceso(1).Caption = Choose(gsIdioma, "&Restore de Información", "&Restore of Information")
  chkProceso.Caption = Choose(gsIdioma, "Restore General", "Restore General")
  lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Información...", "Processing Information...")
  lblProgreso(1).Caption = Choose(gsIdioma, "Procesando Archivo...", "Processing File...")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']

   With porstCodro
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodDro, "
      .Source = .Source & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
      .Source = .Source & "FROM CoDro "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With

End Sub

Private Sub optProceso_Click(Index As Integer)

chkProceso.Enabled = (Index = 1)
frmTablas.Enabled = (Index = 0)
frmTransacciones.Enabled = (Index = 0)
If Index = 1 Then Call chkProceso_Click

End Sub
Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                          'Cambiar (añadir índices).
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + frmTransacciones.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + frmTransacciones.Left + txtDato(tnIndex).Left
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
      With porstCodro
         .MoveFirst
         .Find "CodDro='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
         End If
      End With
   End Select
End Function
