VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B9D22273-0C24-101B-AEBD-04021C009402}#1.0#0"; "KeySta32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.MDIForm fMenu 
   BackColor       =   &H00808080&
   Caption         =   "Sistema - "
   ClientHeight    =   5265
   ClientLeft      =   1110
   ClientTop       =   2175
   ClientWidth     =   8880
   Icon            =   "menu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.PictureBox pctEscritorio 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   8880
      TabIndex        =   26
      Top             =   975
      Visible         =   0   'False
      Width           =   8880
   End
   Begin Threed.SSPanel pan3D1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   4950
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   556
      _StockProps     =   15
      ForeColor       =   255
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      ShadowColor     =   1
      Begin Threed.SSPanel panTime 
         Height          =   315
         Left            =   7605
         TabIndex        =   13
         Top             =   0
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "99:99 am/pm"
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         ShadowColor     =   1
      End
      Begin Threed.SSPanel panDate 
         Height          =   315
         Left            =   8430
         TabIndex        =   14
         Top             =   0
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "dd/mm/aaaa"
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         ShadowColor     =   1
      End
      Begin Threed.SSPanel panKeys 
         Height          =   315
         Index           =   0
         Left            =   9510
         TabIndex        =   15
         Top             =   0
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "NUM"
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         ShadowColor     =   1
      End
      Begin Threed.SSPanel panKeys 
         Height          =   315
         Index           =   1
         Left            =   10140
         TabIndex        =   16
         Top             =   0
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "CAPS"
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         ShadowColor     =   1
      End
      Begin Threed.SSPanel panKeys 
         Height          =   315
         Index           =   2
         Left            =   10755
         TabIndex        =   17
         Top             =   0
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "INS"
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         ShadowColor     =   1
      End
      Begin Threed.SSPanel panKeys 
         Height          =   315
         Index           =   3
         Left            =   11370
         TabIndex        =   18
         Top             =   0
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "SCR"
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         ShadowColor     =   1
      End
      Begin Threed.SSPanel panMessage 
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   7590
         _Version        =   65536
         _ExtentX        =   13388
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   16711680
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         Alignment       =   1
         Begin Threed.SSPanel panPercent 
            Height          =   315
            Left            =   1770
            TabIndex        =   25
            Top             =   0
            Visible         =   0   'False
            Width           =   5820
            _Version        =   65536
            _ExtentX        =   10266
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
            FloodType       =   1
         End
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   820
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   1
      RoundedCorners  =   0   'False
      ShadowColor     =   1
      Font3D          =   1
      Begin MSMAPI.MAPIMessages mpmMensaje 
         Left            =   2805
         Top             =   -60
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession mpsSesion 
         Left            =   2205
         Top             =   -75
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin Crystal.CrystalReport CryReport 
         Left            =   1560
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrinterName     =   "PDF COMPLETE"
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowProgressCtls=   0   'False
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSComDlg.CommonDialog cdlDialogo 
         Left            =   840
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer timer 
         Left            =   30
         Top             =   0
      End
      Begin KeyStatLib.MhState keyStat 
         Height          =   420
         Index           =   3
         Left            =   5505
         TabIndex        =   23
         Top             =   15
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   65
         BackColor       =   4210688
         Style           =   3
         Autosize        =   -1  'True
         TimerInterval   =   1
         MouseIcon       =   "menu.frx":5144A
      End
      Begin KeyStatLib.MhState keyStat 
         Height          =   420
         Index           =   2
         Left            =   4830
         TabIndex        =   22
         Top             =   15
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   65
         BackColor       =   4210688
         Style           =   2
         Autosize        =   -1  'True
         TimerInterval   =   1
         MouseIcon       =   "menu.frx":51466
      End
      Begin KeyStatLib.MhState keyStat 
         Height          =   420
         Index           =   1
         Left            =   4215
         TabIndex        =   21
         Top             =   15
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   65
         BackColor       =   4210688
         Autosize        =   -1  'True
         TimerInterval   =   1
         MouseIcon       =   "menu.frx":51482
      End
      Begin KeyStatLib.MhState keyStat 
         Height          =   420
         Index           =   0
         Left            =   3615
         TabIndex        =   20
         Top             =   15
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   65
         BackColor       =   4210688
         Style           =   1
         Autosize        =   -1  'True
         TimerInterval   =   1
         MouseIcon       =   "menu.frx":5149E
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   465
      Visible         =   0   'False
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Begin VB.ComboBox cmbEjercicio 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   90
         Width           =   2925
      End
      Begin Threed.SSCommand cmdAcercaDe 
         Height          =   360
         Left            =   9660
         TabIndex        =   1
         ToolTipText     =   "Información del Sistema"
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":514BA
      End
      Begin Threed.SSCommand cmdSalir 
         Height          =   360
         Left            =   11250
         TabIndex        =   3
         ToolTipText     =   "Finalizar Sistema"
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":514D6
      End
      Begin Threed.SSCommand cmdAyuda 
         Height          =   360
         Left            =   10250
         TabIndex        =   2
         ToolTipText     =   "Ayuda del Sistema"
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":514F2
      End
      Begin Threed.SSCommand cmdPersonal 
         Height          =   360
         Left            =   270
         TabIndex        =   6
         Top             =   90
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":5150E
      End
      Begin Threed.SSCommand cmdFormula 
         Height          =   360
         Left            =   840
         TabIndex        =   7
         Top             =   90
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":5152A
      End
      Begin Threed.SSCommand cmdAsistencia 
         Height          =   360
         Left            =   1410
         TabIndex        =   8
         Top             =   90
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":51546
      End
      Begin Threed.SSRibbon ribMoneda 
         Height          =   360
         Index           =   1
         Left            =   6765
         TabIndex        =   11
         Top             =   90
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "menu.frx":51562
      End
      Begin Threed.SSRibbon ribMoneda 
         Height          =   360
         Index           =   0
         Left            =   6360
         TabIndex        =   10
         Top             =   90
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "menu.frx":5157E
      End
      Begin Threed.SSCommand cmdEmpresa 
         Height          =   360
         Index           =   0
         Left            =   8085
         TabIndex        =   4
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":5159A
      End
      Begin Threed.SSCommand cmdEmpresa 
         Height          =   360
         Index           =   1
         Left            =   8505
         TabIndex        =   5
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "menu.frx":515B6
      End
   End
   Begin VB.Menu mnuOpcion1 
      Caption         =   ""
      WindowList      =   -1  'True
      Begin VB.Menu mnuOpcion10 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOpcion2 
      Caption         =   ""
      Begin VB.Menu mnuOpcion20 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOpcion3 
      Caption         =   ""
      Begin VB.Menu mnuOpcion30 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOpcion4 
      Caption         =   ""
      Begin VB.Menu mnuOpcion40 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOpcion5 
      Caption         =   ""
      Begin VB.Menu mnuOpcion50 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOpcion6 
      Caption         =   ""
      Begin VB.Menu mnuOpcion60 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOpcion7 
      Caption         =   ""
   End
End
Attribute VB_Name = "fMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                     ' Declarar variable antes de usarla

Private s_OldMessage As String                      ' Mensaje temporal e proceso principal del
Private Sub cmbEjercicio_Click()
  ' Actualizo la variable del ejercicio
  ps_Anyo = Right$(cmbejercicio.Text, 4)
  ps_DaBasCon = IIf(Left(ps_DaBasCon, 1) = "c", "c" & ps_CodEmpresa & ps_Anyo, ps_DataBase)
  
  s_Sql = "SELECT count(column_name) as valor FROM information_schema.COLUMNS WHERE table_name = 'plcfgempresa' and column_name='fecha_limiteproc' and LTRIM(table_schema)='" & ps_DataBase & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If porstRecordset!valor = 0 Then
    Flag_RestringeSistema = "SIN RESTRICCIONES"
  Else
    Flag_RestringeSistema = "RESTRINGIR"
  End If
  
  If Flag_RestringeSistema = "RESTRINGIR" Then
    s_Sql = "SELECT nivelcencosto, fecha_limiteproc FROM plcfgempresa WHERE pdoano='" & ps_Anyo & "'"
  Else
    s_Sql = "SELECT nivelcencosto FROM plcfgempresa WHERE pdoano='" & ps_Anyo & "'"
  End If
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Nivel de centro de costo
  pn_NivelCenCosto = 5
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    pn_NivelCenCosto = CInt(porstRecordset!nivelcencosto)
    ' Inicializo fecha limite restringir algunos procesos del sistema
    If Flag_RestringeSistema = "RESTRINGIR" Then
      ps_Fecha_LimiteProc = IIf(IsNull(porstRecordset!fecha_limiteproc) = False, gdl_Funcion.Desencripta(porstRecordset!fecha_limiteproc), Format(Date, "dd-mm-yyyy"))
    End If
  End If
  Set porstRecordset = Nothing
  
End Sub
Private Sub cmdAcercaDe_Click()
  fAcercaDe.Show vbModal
End Sub
Private Sub cmdAsistencia_Click()
  s_SwRegistro = "asistencia": LoadOpcion o_SelAsistencia, "2", 5, n_FormatoReg
End Sub
Private Sub cmdEmpresa_Click(Index As Integer)
  Dim s_CodEmpresa As String, s_NomEmpresa As String, s_RucEmpresa As String
  Dim s_EmpresaCon As String, s_DataBase As String, s_DaBasCon As String
  Dim s_ClsPlanilla As String, s_DesClsPlanilla As String
  
  ' Verifico que no exista ventanas activas
  If Forms.Count >= 2 Then Beep: MsgBox "Debe cerra las pantallas activas : '" & Forms(1).Caption & "'", vbExclamation: Exit Sub
  If Index = 0 Then  ' Selcción de empresa
    ' Guardo los valores actuales
    s_CodEmpresa = ps_CodEmpresa: s_NomEmpresa = ps_NomEmpresa
    s_RucEmpresa = ps_RucEmpresa: s_DataBase = ps_DataBase
    s_EmpresaCon = ps_EmpresaCon: s_DaBasCon = ps_DaBasCon
    s_ClsPlanilla = ps_ClsPlanilla: s_DesClsPlanilla = ps_DesClsPlanilla
    fSelEmpresa.Show vbModal
    ' Restablesco los valores
    If Not pl_Salir Then
      ps_CodEmpresa = s_CodEmpresa: ps_NomEmpresa = s_NomEmpresa
      ps_RucEmpresa = s_RucEmpresa: ps_DataBase = s_DataBase
      ps_EmpresaCon = s_EmpresaCon: ps_DaBasCon = s_DaBasCon
      ps_ClsPlanilla = s_ClsPlanilla: ps_DesClsPlanilla = s_DesClsPlanilla
    End If
  Else  ' Selección de clase planilla
    s_ClsPlanilla = ps_ClsPlanilla
    s_DesClsPlanilla = ps_DesClsPlanilla
    fSelPlanilla.Show vbModal
    ' Restablesco los valores
    If Not pl_Salir Then
      ps_ClsPlanilla = s_ClsPlanilla: ps_DesClsPlanilla = s_DesClsPlanilla
    End If
  End If
  ' Nombre del Modulo del sistema
  fMenu.Caption = ps_NomSistema & " - " & ps_NomEmpresa & " - " & ps_DesClsPlanilla

End Sub
Private Sub cmdFormula_Click()
  LoadOpcion fFormulaConcepto, "1", 2, n_FormatoReg
End Sub
Private Sub cmdPersonal_Click()
  LoadOpcion fPersonal, "2", 1, n_FormatoReg
End Sub
Private Sub cmdSalir_Click()
  Unload Me
End Sub
Private Sub keyStat_Change(Index As Integer)
  ' Muestro la Tecla Activa
  panKeys(Index).Font.Bold = Not panKeys(Index).Font.Bold
End Sub
Private Sub MDIForm_Load()

  Dim n_Op1 As Byte, n_Op2 As Byte, n_Op3 As Byte, n_Op4 As Byte, n_Op5 As Byte, n_Op6 As Byte
  Dim s_Archivo As String, n_Index As Integer
  
  Me.WindowState = vbMaximized
 
  fMenu.Caption = ps_NomSistema & " - " & ps_Licencia
  ' Verifico que exista el Icono del Sistema y Papel Tapiz
  fMenu.Icon = LoadPicture()
  's_Archivo = gdl_Procedure.ps_PathImagen & "planilla.ico"
  If dir$(s_Archivo, vbNormal) <> "" Then
     fMenu.Icon = LoadPicture(s_Archivo)
  End If

  ' Fondo de escritorio
  pctEscritorio.Visible = False
  pctEscritorio.AutoRedraw = True
  pctEscritorio.Picture = LoadPicture()
  s_Archivo = gdl_Procedure.ps_PathImagen & "planilla.jpg"
  If dir$(s_Archivo, vbNormal) <> "" Then
     pctEscritorio.Picture = LoadPicture(s_Archivo)
  End If
  pctEscritorio.Refresh

  ' Cargo las Variables Publicas del Sistema
  gs_FechaHora = Now
  panTime.Caption = Format$(gs_FechaHora, s_FormatoHora_1)
  panDate.Caption = Format(gs_FechaHora, s_FormatoFecha)
  panKeys(0).Font.Bold = keyStat(0).Value
  panKeys(1).Font.Bold = keyStat(1).Value
  panKeys(2).Font.Bold = keyStat(2).Value
  panKeys(3).Font.Bold = keyStat(3).Value
    
  ' Fuerzo a Mostrar el Menu
  timer.Interval = 1
  Show
  ' Nombre del Modulo del sistema
  fMenu.Caption = ps_NomSistema & " - " & ps_NomEmpresa & " - " & ps_DesClsPlanilla
  panToolBar.Visible = True
    
  ' Cambio el Puntero a Espera
  gdl_Procedure.PunteroEnEspera
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Cargando Opciones..."
  panPercent.Visible = True
    
  ' Cargo los Periodos del Ejercicio Activo
  For n_Index = (Val(ps_Anyo) - 5) To (Val(ps_Anyo) + 5)
      cmbejercicio.AddItem "Proceso " & n_Index
  Next n_Index
  cmbejercicio.ListIndex = 5
    
  ' Cargo los datos del menu
  s_Sql = "SELECT mnu.opcion, mnu.orden, mnu.detmdl, opc.codusr"
  s_Sql = s_Sql & " FROM sgmdl mnu INNER JOIN sgpms opc ON mnu.codsis=opc.codsis AND mnu.codmdl=opc.codmdl"
  s_Sql = s_Sql & " WHERE mnu.codsis='" & ps_CodSistema & "'"
  s_Sql = s_Sql & " AND opc.codemp='" & ps_CodEmpresa & "'"
  s_Sql = s_Sql & " AND opc.codusr='" & ps_Usuario & "'"
  s_Sql = s_Sql & " UNION"
  s_Sql = s_Sql & " SELECT mnu.opcion, mnu.orden, mnu.detmdl, null"
  s_Sql = s_Sql & " FROM sgmdl mnu"
  s_Sql = s_Sql & " WHERE mnu.codsis='" & ps_CodSistema & "'"
  s_Sql = s_Sql & " AND NOT EXISTS(SELECT * FROM sgpms opc"
  s_Sql = s_Sql & " WHERE opc.codemp='" & ps_CodEmpresa & "'"
  s_Sql = s_Sql & " AND opc.codusr='" & ps_Usuario & "'"
  s_Sql = s_Sql & " AND opc.codmdl= mnu.codmdl"
  s_Sql = s_Sql & " AND opc.codsis = mnu.codsis)"
  s_Sql = s_Sql & " ORDER BY opcion, orden"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
  
  n_Op1 = 0: n_Op2 = 0: n_Op3 = 0: n_Op4 = 0: n_Op5 = 0: n_Op6 = 0
  'Carga todas la opciones del sistema en arreglos
  While Not porstRecordset.EOF
    Select Case porstRecordset!opcion
     Case "0"
      Select Case porstRecordset!orden
       Case "01": mnuOpcion1.Caption = porstRecordset!detmdl
       Case "02": mnuOpcion2.Caption = porstRecordset!detmdl
       Case "03": mnuOpcion3.Caption = porstRecordset!detmdl
       Case "04": mnuOpcion4.Caption = porstRecordset!detmdl
       Case "05": mnuOpcion5.Caption = porstRecordset!detmdl
       Case "06": mnuOpcion6.Caption = porstRecordset!detmdl
       Case "07": mnuOpcion7.Caption = porstRecordset!detmdl
      End Select
     Case "1"
      n_Op1 = n_Op1 + 1
      Load mnuOpcion10(n_Op1)
      mnuOpcion10(n_Op1).Caption = porstRecordset!detmdl
     Case "2"
      n_Op2 = n_Op2 + 1
      Load mnuOpcion20(n_Op2)
      mnuOpcion20(n_Op2).Caption = porstRecordset!detmdl
     Case "3"
      n_Op3 = n_Op3 + 1
      Load mnuOpcion30(n_Op3)
      mnuOpcion30(n_Op3).Caption = porstRecordset!detmdl
     Case "4"
      n_Op4 = n_Op4 + 1
      Load mnuOpcion40(n_Op4)
      mnuOpcion40(n_Op4).Caption = porstRecordset!detmdl
     Case "5"
      n_Op5 = n_Op5 + 1
      Load mnuOpcion50(n_Op5)
      mnuOpcion50(n_Op5).Caption = porstRecordset!detmdl
     Case "6"
      n_Op6 = n_Op6 + 1
      Load mnuOpcion60(n_Op6)
      mnuOpcion60(n_Op6).Caption = porstRecordset!detmdl
    End Select
    panPercent.FloodPercent = ((porstRecordset.AbsolutePosition) * 100) \ porstRecordset.RecordCount
    DoEvents
    porstRecordset.MoveNext
  Wend
  ' Cierro el recordset y saco del entorno
  porstRecordset.Close: Set porstRecordset = Nothing
  
  ' Visualizo las opciones del menu
  mnuOpcion10(0).Visible = (n_Op1 > 0)
  mnuOpcion20(0).Visible = (n_Op2 > 0)
  mnuOpcion30(0).Visible = (n_Op3 > 0)
  mnuOpcion40(0).Visible = (n_Op4 > 0)
  mnuOpcion50(0).Visible = (n_Op5 > 0)
  mnuOpcion60(0).Visible = (n_Op6 > 0)
    
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
    
  ' Cargo los graficos del botón de monedas
  For n_Index = 0 To 1
    ribMoneda(n_Index).PictureUp = LoadPicture()
    ribMoneda(n_Index).ToolTipText = "Analisis en " & Choose(n_Index + 1, s_Codmon_mn_Nom, s_Codmon_me_Nom)
    s_Archivo = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "soles", "dolares") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Archivo) Then ribMoneda(n_Index).PictureUp = LoadPicture(s_Archivo)
  Next n_Index
  ribMoneda(0).Value = True
    
  ' Cargo los graficos del botón de acerca, ayuda, salida
  gdl_Procedure.LoadGrafics cmdPersonal, "asiperso", "Registro de Personal"
  gdl_Procedure.LoadGrafics cmdFormula, "formula", "Definición de Formulas"
  gdl_Procedure.LoadGrafics cmdAsistencia, "promedio", "Registro de Asistencia"
  gdl_Procedure.LoadGrafics cmdEmpresa(0), "impempre", "Selección de Empresa"
  gdl_Procedure.LoadGrafics cmdEmpresa(1), "impconso", "Selección de Clase Planilla"
  gdl_Procedure.LoadGrafics cmdAcercaDe, "acercade", ""
  gdl_Procedure.LoadGrafics cmdAyuda, "ayuda", ""
  gdl_Procedure.LoadGrafics cmdSalir, "escapar", ""
    
  ' Cambio el Puntero a Normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub MDIForm_Resize()

  Dim ImageWidth As Single
  Dim ImageHeight As Single
  Dim Red As Double
  
  On Error Resume Next
  pctEscritorio.Height = Me.ScaleHeight
  ' Altura imagen es mayor que escritorio
  If (pctEscritorio.Picture.Height > fMenu.Height) Then
    Red = (fMenu.Height / pctEscritorio.Picture.Height)
    ImageHeight = (pctEscritorio.Picture.Height * Red)
    ImageWidth = (pctEscritorio.Picture.Width * Red)
  End If
  ' Ancho imagen es mayor que escritorio
  If (ImageWidth > fMenu.Width) Then
    Red = (fMenu.Width / fMenu.pctEscritorio.Picture.Width)
    ImageHeight = pctEscritorio.Picture.Height * Red
    ImageWidth = pctEscritorio.Picture.Width * Red
  End If
  
  pctEscritorio.PaintPicture pctEscritorio.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0
  pctEscritorio.PaintPicture pctEscritorio.Picture, 0, 0, ImageWidth, ImageHeight, 0, 0
  
  ' imagen de escritorio
  Set fMenu.Picture = pctEscritorio.Image
  
  Resume

End Sub
Private Sub mnuOpcion10_Click(Index As Integer)

  Select Case Index
   Case 1: LoadOpcion fConceptoPlanilla, "1", Index, n_FormatoReg
   Case 2: LoadOpcion fFormulaConcepto, "1", Index, n_FormatoReg
   ' 3 - Linea
   Case 4: LoadOpcion fPeriodoPago, "1", Index, n_FormatoReg
   Case 5: LoadOpcion fCentroCosto, "1", Index, n_FormatoReg
   Case 6: LoadOpcion fProcesoCalculo, "1", Index, n_FormatoReg
   ' 7 - Linea
   Case 8: LoadOpcion fBoletaPago, "1", Index, n_FormatoReg
   Case 9: LoadOpcion fGeneradoReporte, "1", Index, n_FormatoReg
   ' 10 - Linea
   Case 11: LoadOpcion fTablasGeneral, "1", Index, n_FormatoReg
   Case 12: LoadOpcion fEntidadPension, "1", Index, n_FormatoReg
   ' 13 Linea
   Case 14: LoadOpcion fEstablecimientoPropio, "1", Index, n_FormatoReg
   Case 15: LoadOpcion fEmpresasqdes, "1", Index, n_FormatoReg
   Case 16: LoadOpcion fEmpresaseqdes, "1", Index, n_FormatoReg
   Case 17: LoadOpcion fEmpresasqmdes, "1", Index, n_FormatoReg
  End Select

End Sub
Private Sub mnuOpcion20_Click(Index As Integer)
  
  Select Case Index
   Case 1: LoadOpcion fPersonal, "2", Index, n_FormatoReg
   Case 2: LoadOpcion fSelPersonal, "2", Index, n_FormatoReg
   ' 3 - Linea
   Case 4: s_SwRegistro = "exepcional": LoadOpcion o_SelExepcional, "2", Index, n_FormatoReg
   Case 5: s_SwRegistro = "asistencia": LoadOpcion o_SelAsistencia, "2", Index, n_FormatoReg
   Case 6: s_SwRegistro = "distcencos": LoadOpcion o_SelDisCenCosto, "2", Index, n_FormatoReg
   ' 7 - Linea
   Case 8: s_SwRegistro = "calpllgral": LoadOpcion fCalculoPlanilla, "2", Index, n_FormatoPrc, vbModal
   Case 9: s_SwRegistro = "calpllpers": LoadOpcion o_CalculoPersona, "2", Index, n_FormatoPrc
   Case 10: s_SwRegistro = "inipllgral": LoadOpcion o_DepuraCalculo, "2", Index, n_FormatoPrc, vbModal
   ' 11 - Linea
   Case 12: s_SwRegistro = "pllasconta": LoadOpcion o_ContaPlanilla, "2", Index, n_FormatoPrc
   Case 13: s_SwRegistro = "rpcontapla": LoadOpcion o_RepContaPlani, "2", Index, n_FormatoRpt
  End Select

End Sub
Private Sub mnuOpcion30_Click(Index As Integer)
  
  Select Case Index
   Case 1: LoadOpcion fEscalaQuinta, "3", Index, n_FormatoCst
   Case 2: s_SwRegistro = "anrenta5ta": LoadOpcion o_SelRentaQuinta, "3", Index, n_FormatoRpt
   ' 3 - Linea
   Case 4: s_SwRegistro = "consulxcpc": LoadOpcion o_SelConsulxcpc, "3", Index, n_FormatoCst
   Case 5: s_SwRegistro = "consulxpsn": LoadOpcion o_SelConsulxpsn, "3", Index, n_FormatoCst
   ' 6 - Linea
   Case 7: s_SwRegistro = "anvacacion": LoadOpcion o_SelVacacionAna, "3", Index, n_FormatoCst
   Case 8: LoadOpcion fConsultaVarios, "3", Index, n_FormatoCst
  End Select

End Sub
Private Sub mnuOpcion40_Click(Index As Integer)
  Select Case Index
   Case 1: s_SwRegistro = "pvsvacacio": LoadOpcion o_PvsVacaciones, "4", Index, n_FormatoPvs
   Case 2: s_SwRegistro = "pvsgratifi": LoadOpcion o_PvsGratifica, "4", Index, n_FormatoPvs
   ' 3 - Linea
   Case 4: s_SwRegistro = "pvscoxtise": LoadOpcion o_PvsComxTieSer, "4", Index, n_FormatoPvs
   Case 5: s_SwRegistro = "repcoxtise": LoadOpcion o_RepComxTieSer, "4", Index, n_FormatoRpt
   ' 6 - Linea
   Case 7: s_SwRegistro = "pvscontabi": LoadOpcion o_ContaProvision, "4", Index, n_FormatoPrc
  End Select
End Sub
Private Sub mnuOpcion50_Click(Index As Integer)

  Select Case Index
   Case 1: LoadOpcion fReporteBoleta, "5", Index, n_FormatoRpt
   Case 2: s_SwRegistro = "recibopago": LoadOpcion o_RptReciboPago, "5", Index, n_FormatoRpt
   Case 3: s_SwRegistro = "repliquida": LoadOpcion o_RepLiquidacion, "5", Index, n_FormatoRpt
   Case 4: LoadOpcion o_SelReporGnral, "5", Index, n_FormatoRpt
   ' 5 - Linea
   Case 6: s_SwRegistro = "planitraba": LoadOpcion o_RepPrePlanilla, "5", Index, n_FormatoRpt
   Case 7: LoadOpcion o_PlanillaGnral, "5", Index, n_FormatoRpt
   Case 8: LoadOpcion fReporPlanillAfp, "5", Index, n_FormatoRpt
   ' 9 - Linea
   Case 10: s_SwRegistro = "expinfopdt": LoadOpcion o_ExportarSunat, "5", Index, n_FormatoRpt
   Case 11: LoadOpcion fTransferBancos, "5", Index, n_FormatoRpt
   ' 12 - Linea
   Case 13: s_SwRegistro = "certifi5ta": LoadOpcion o_Certifikdo5ta, "5", Index, n_FormatoRpt
   Case 14: s_SwRegistro = "certifisnp": LoadOpcion o_CertifikdoSnp, "5", Index, n_FormatoRpt
   Case 15: s_SwRegistro = "certifiafp": LoadOpcion o_CertifikdoAfp, "5", Index, n_FormatoRpt
   Case 16: s_SwRegistro = "certifiuti": LoadOpcion o_CertifikdoUti, "5", Index, n_FormatoRpt
   ' 17 - Linea
   Case 18: s_SwRegistro = "disbillete": LoadOpcion o_RptDisBillete, "5", Index, n_FormatoRpt
  End Select
End Sub
Private Sub mnuOpcion60_Click(Index As Integer)
  Select Case Index
   Case 1: LoadOpcion fTablaSistema, "6", Index, n_FormatoReg
   Case 2: LoadOpcion fEmpresa, "6", Index, n_FormatoReg
   ' 3 - Linea
   Case 4: LoadOpcion fUsuario, "6", Index, n_FormatoReg
   Case 5: LoadOpcion fSeguridad, "6", Index, n_FormatoLbr
   Case 6: LoadOpcion fCambioPassword, "6", Index, n_FormatoLbr
   Case 7: LoadOpcion fCambioPassword, "6", Index, n_FormatoRpt
   ' 8 - Linea
   Case 9: LoadOpcion fTransInformacio, "6", Index, n_FormatoPrc
   Case 10: LoadOpcion fAbcBckRest, "6", Index, n_FormatoPrc
   ' 11 - Linea
   Case 12: LoadOpcion fExporTRegistro, "6", Index, n_FormatoPrc
   Case 13: LoadOpcion fExportSunat, "6", Index, n_FormatoPrc, vbModal
   ' 14 - Linea
   Case 15: LoadOpcion fAcercaDe, "6", Index, n_FormatoLbr, vbModal
  End Select
End Sub
Private Sub mnuOpcion7_Click()
  End
End Sub
Private Sub panKeys_Click(Index As Integer)
  panKeys(Index).Font.Bold = Not panKeys(Index).Font.Bold
  keyStat(Index).Value = Not keyStat(Index).Value
End Sub
' Private Sub Timer_Timer()
'  ' Muestro la Hora en la Barra Inferior
'  panTime.Caption = Format$(gs_FechaHora, s_FormatoHora_1)
'  DoEvents
'End Sub
