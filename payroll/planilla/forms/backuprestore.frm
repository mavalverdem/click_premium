VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOutl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fBackupRestore 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5520
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7320
   Icon            =   "backuprestore.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7320
   Begin Threed.SSFrame sfmProgreso 
      Height          =   480
      Left            =   75
      TabIndex        =   21
      Top             =   5535
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   " Procesando archivo : "
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
         Left            =   45
         TabIndex        =   22
         Top             =   225
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   4860
      Left            =   75
      TabIndex        =   0
      Top             =   600
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   8573
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabHeight       =   520
      TabMaxWidth     =   3052
      BackColor       =   -2147483644
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Backup/Restore"
      TabPicture(0)   =   "backuprestore.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "shpCuadro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmCuadro(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCuadro(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbPeriodo(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbPeriodo(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbEjercicio(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbEjercicio(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.ComboBox cmbEjercicio 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         ItemData        =   "backuprestore.frx":0028
         Left            =   540
         List            =   "backuprestore.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4100
         Width           =   1050
      End
      Begin VB.ComboBox cmbEjercicio 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         ItemData        =   "backuprestore.frx":002C
         Left            =   3930
         List            =   "backuprestore.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4100
         Width           =   1050
      End
      Begin VB.ComboBox cmbPeriodo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         ItemData        =   "backuprestore.frx":0030
         Left            =   5025
         List            =   "backuprestore.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   4100
         Width           =   1695
      End
      Begin VB.ComboBox cmbPeriodo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         ItemData        =   "backuprestore.frx":0034
         Left            =   1635
         List            =   "backuprestore.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4100
         Width           =   1695
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2490
         Index           =   1
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   4392
         _StockProps     =   14
         Caption         =   " Ubicación "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   2235
         End
         Begin VB.DriveListBox drbUnidad 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2240
         End
         Begin VB.Label lblDato 
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   255
            Width           =   1005
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1095
         Index           =   2
         Left            =   4560
         TabIndex        =   7
         Top             =   2655
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   1931
         _StockProps     =   14
         Caption         =   " Opción "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin Threed.SSOption optParametro 
            Height          =   200
            Index           =   0
            Left            =   230
            TabIndex        =   8
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Backup"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optParametro 
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   9
            Top             =   525
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Restore"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCheck chkParametro 
            Height          =   210
            Left            =   225
            TabIndex        =   10
            Top             =   780
            Width           =   1950
            _Version        =   65536
            _ExtentX        =   3440
            _ExtentY        =   370
            _StockProps     =   78
            Caption         =   "Opción General"
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
            Font3D          =   1
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3660
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   6456
         _StockProps     =   14
         Caption         =   " Información a Procesar "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin MSOutl.Outline outParametro 
            Height          =   3195
            Left            =   255
            TabIndex        =   2
            Top             =   330
            Width           =   3840
            _Version        =   65536
            _ExtentX        =   6773
            _ExtentY        =   5636
            _StockProps     =   77
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
            MouseIcon       =   "backuprestore.frx":0038
            PicturePlus     =   "backuprestore.frx":0054
            PictureMinus    =   "backuprestore.frx":014E
            PictureLeaf     =   "backuprestore.frx":0248
            PictureOpen     =   "backuprestore.frx":076A
            PictureClosed   =   "backuprestore.frx":0864
         End
         Begin VB.Image imgGrafico 
            Height          =   240
            Left            =   75
            Top             =   3285
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   3345
            Index           =   1
            Left            =   45
            Shape           =   4  'Rounded Rectangle
            Top             =   270
            Width           =   4260
         End
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Final :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   3810
         TabIndex        =   14
         Top             =   3810
         Width           =   1125
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo Inicial :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   585
         TabIndex        =   11
         Top             =   3870
         Width           =   1125
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   705
         Index           =   0
         Left            =   105
         Shape           =   4  'Rounded Rectangle
         Top             =   3780
         Width           =   6945
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7320
      _Version        =   65536
      _ExtentX        =   12912
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
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   6435
         TabIndex        =   19
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "backuprestore.frx":0D86
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6045
         TabIndex        =   20
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "backuprestore.frx":0DA2
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   450
         TabIndex        =   18
         Top             =   120
         Width           =   5085
      End
   End
End
Attribute VB_Name = "fBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private s_Registro As String                            ' Codigo del registro
'[
Private Sub ppBackup_Informacion()
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Archivo As String
  Dim nRegistro As Long, nRegistros As Long
  Dim sAñoIni As String, sAñoFin As String, sMesIni As String, sMesFin As String
  Dim nSecuencia As Integer, nContador As Integer
  Dim nElemento As Integer, nPicture As Integer
  Dim a_Tabla(), a_Archivo(), a_Columnas(), a_Where()
  
  ' Cambio el Mensaje y Muestro la Barra
  pgbProgreso.Value = pgbProgreso.Min
  
  sAñoIni = Left(Trim(cmbejercicio(0).Text), 4)
  sAñoFin = Left(Trim(cmbejercicio(1).Text), 4)
  sMesIni = Left(Trim(cmbPeriodo(0).Text), 2)
  sMesFin = Left(Trim(cmbPeriodo(1).Text), 2)
  ' Creo objeto de archivo
  Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
  ' Exporto la información seleccionada
  nContador = 1: nElemento = 0
  While outParametro.ListCount > nContador
    fMenu.panPercent.Visible = True
    ' Nivel de proceso o selección de información
    If outParametro.Indent(nContador) <> 1 Then
      ' Selecciono el nombre y columnas del archivo de texto
      a_Archivo = Choose(nElemento + 1, Array("cta"), Array("cfg"), Array("cco"), Array("tcb"), _
          Array("bco"), Array("dci"), Array("tpt"), Array("pfs"), Array("dsm"), Array("via"), Array("zon"), _
          Array("cpc"), Array("cls"), Array("afp"), Array("eps"), Array("bas"), Array("cgo"), Array("ubi"), Array("sec"), _
          Array("var"), Array("esq"), Array("pem"), Array("pck"), _
          Array("prc"), Array("cxp"), Array("cpr"), Array("cxc"), Array("cpv"), _
          Array("psn"), Array("rxd"), Array("con"), Array("esr"), Array("exl"), Array("dfm"), _
          Array("boc", "bod"), Array("rpc", "rpd"), Array("rpl", "pld"), _
          Array("pdo"), Array("cte"), Array("asi"), Array("rde"), _
          Array("rhc", "rda"), Array("car"), Array("paf"), _
          Array("vae", "vac", "vad"), Array("gre", "gra"), Array("cse", "csm", "csd", "csc"))
      'a_Tabla = Choose(nElemento + 1, Array("cocta"), Array("cocfg"), Array("cocco"), Array("tgtcb"),
      a_Tabla = Choose(nElemento + 1, Array("cocta"), Array("cocfg"), Array("cocco"), _
          Array("plbanco"), Array("pldocidentidad"), Array("pltpotrabajador"), Array("plprofesion"), Array("pldstmoneda"), Array("pltipovia"), Array("pltipozona"), _
          Array("plconcepto"), Array("plclasplan"), Array("plentidadafp"), Array("plentidadeps"), Array("pltablabase"), Array("plcargo"), Array("plubicacion"), Array("plseccion"), _
          Array("plvarfunc"), Array("plescalaquinta"), Array("plcfgempresa"), Array("plparametroafp"), _
          Array("plproceso"), Array("plconceplanilla"), Array("plconceproceso"), Array("plctacencos"), Array("plctapvs"), _
          Array("plpersonal"), Array("plremudefa"), Array("plcontrato"), Array("plestudios"), Array("plexpelaboral"), Array("plfamiliares"), _
          Array("plboletapago", "pldetaboleta"), Array("plgenreporte", "pldetareporte"), Array("plplanilla", "pldetaplanilla"), _
          Array("plperiodo"), Array("plcuentacte"), Array("plasistencia"), Array("plremuexce"), _
          Array("plresultado", "pldatoresultado"), Array("plcartabanco"), Array("plplanillafp"), _
          Array("plpvsperiodovac", "plpvsvacacion", "plpvsvacaciondet"), Array("plpvsperiodogra", "plpvsgratifica"), Array("plctsperiodo", "plctsperiodosub", "plctsmovimiento", "plctsresultado"))
      a_Where = Choose(nElemento + 1, Array(""), Array(""), Array(""), Array("WHERE DATE_FORMAT(tcb.fehtcb, '%Y%m')>='" & sAñoIni & sMesIni & "' AND DATE_FORMAT(tcb.fehtcb, '%Y%m')<='" & sAñoFin & sMesFin & "'"), _
          Array(""), Array(""), Array(""), Array(""), Array(""), Array(""), Array(""), _
          Array(""), Array(""), Array(""), Array(""), Array("WHERE bas.pdoano>='" & sAñoIni & "' AND bas.pdoano<='" & sAñoFin & "'"), Array(""), Array(""), Array(""), _
          Array(""), Array(""), Array("WHERE pem.pdoano>='" & sAñoIni & "' AND pem.pdoano>='" & sAñoFin & "'"), Array("WHERE pck.pdoano>='" & sAñoIni & "' AND pck.pdoano>='" & sAñoFin & "'"), _
          Array(""), Array(""), Array(""), Array(""), Array(""), _
          Array("WHERE DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=con.codcls AND psn.codpsn=con.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=esr.codcls AND psn.codpsn=esr.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=exl.codcls AND psn.codpsn=exl.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=dfm.codcls AND psn.codpsn=dfm.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), _
          Array("", ""), Array("", ""), Array("", ""), _
          Array("WHERE CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=cte.codcls AND pdo.codpdo=cte.codpdoprv AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=asi.codcls AND pdo.codpdo=asi.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=rde.codcls AND pdo.codpdo=rde.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE CONCAT(rhc.pdoano, rhc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(rhc.pdoano, rhc.pdomes)<='" & sAñoFin & sMesFin & "'", ", plresultado rhc WHERE rhc.codcls=rda.codcls AND rhc.codpdo=rda.codpdo AND rhc.codpsn=rda.codpsn AND CONCAT(rhc.pdoano, rhc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(rhc.pdoano, rhc.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE DATE_FORMAT(car.fechaproce, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array("WHERE CONCAT(paf.pdoano, paf.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(paf.pdoano, paf.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE vae.codpvs<='" & sAñoFin & "'", "WHERE vac.codpvs<='" & sAñoFin & "'", "WHERE CONCAT(vad.pdoano, vad.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(vad.pdoano, vad.pdomes)<='" & sAñoFin & sMesFin & "'"), Array("WHERE CONCAT(gre.pdoano, gre.mesini)>='" & sAñoIni & sMesIni & "' AND CONCAT(gre.pdoano, gre.mesfin)<='" & sAñoFin & sMesFin & "'", "WHERE CONCAT(gra.pdoano, gra.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(gra.pdoano, gra.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE cse.pdoano>='" & sAñoIni & "' AND cse.pdoano<='" & sAñoFin & "'", ", plctsperiodo cse WHERE cse.codcls=csm.codcls AND cse.pdocts=csm.pdocts AND CONCAT(cse.pdoano, csm.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(cse.pdoano, csm.pdomes)<='" & sAñoFin & sMesFin & "'", ", plctsperiodo cse WHERE cse.codcls=csd.codcls AND cse.pdocts=csd.pdocts AND CONCAT(cse.pdoano, csd.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(cse.pdoano, csd.pdomes)<='" & sAñoFin & sMesFin & "'", "WHERE CONCAT(csc.pdoano, csc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(csc.pdoano, csc.pdomes)<='" & sAñoFin & sMesFin & "'"))
      a_Columnas = Choose(nElemento + 1, Array(26), Array(21), Array(8), Array(7), _
          Array(11), Array(8), Array(7), Array(7), Array(8), Array(8), Array(8), _
          Array(9), Array(10), Array(16), Array(9), Array(22), Array(8), Array(7), Array(7), _
          Array(6), Array(7), Array(49), Array(37), _
          Array(8), Array(11), Array(9), Array(14), Array(15), _
          Array(70), Array(9), Array(15), Array(12), Array(12), Array(29), _
          Array(15, 18), Array(12, 14), Array(19, 9), _
          Array(15), Array(25), Array(38), Array(10), _
          Array(22, 18), Array(16), Array(17), _
          Array(10, 12, 26), Array(12, 25), Array(9, 17, 20, 21))
      nPicture = outParametro.PictureType(nContador)
      ' verifico se encuentra seleccionado
      If nPicture = outClosed Then
        For nSecuencia = 0 To UBound(a_Archivo, 1)
          s_Archivo = dlbDirectorio(0).path & "\" & ps_RucEmpresa & a_Archivo(nSecuencia) & ".bma"
          ' Recupero la información para exportar
          s_Sql = "SELECT " & a_Archivo(nSecuencia) & ".* "
          s_Sql = s_Sql & "FROM " & a_Tabla(nSecuencia) & " " & a_Archivo(nSecuencia) & " "
          s_Sql = s_Sql & IIf(chkParametro.Value, "", a_Where(nSecuencia))
          Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
          If Not (porstRecordset.BOF And porstRecordset.EOF) Then
            nRegistros = porstRecordset.RecordCount: nRegistro = 0
            ' Inicializo la barra de progreso
            pgbProgreso.Max = nRegistros
            pgbProgreso.Value = pgbProgreso.Min
            sfmProgreso.Caption = " Backup Información: " & Trim(outParametro.List(nContador)) & " - " & Right(s_Archivo, 18) & " "
            ' Aperturo el archivo
            Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
            ' Grabo los nombres, tipos de los campos de la tabla seleccionada
            psRegistro = ppRegistro_Texto(porstRecordset, a_Columnas(nSecuencia), 0)
            potxtFileExp.WriteLine psRegistro
            psRegistro = ppRegistro_Texto(porstRecordset, a_Columnas(nSecuencia), 1)
            potxtFileExp.WriteLine psRegistro
            While Not porstRecordset.EOF
              ' Grabo el detalle de tabla
              psRegistro = ppRegistro_Texto(porstRecordset, a_Columnas(nSecuencia), 2)
              potxtFileExp.WriteLine psRegistro
              ' Incremento el porcentaje
              nRegistro = nRegistro + 1
              pgbProgreso.Value = nRegistro
              DoEvents
              porstRecordset.MoveNext
            Wend
            ' Cierro objeto
            potxtFileExp.Close
          End If
        Next nSecuencia
      End If
      nElemento = nElemento + 1
    End If
    fMenu.panPercent.FloodPercent = ((nContador * 100) \ outParametro.ListCount)
    nContador = nContador + 1
  Wend
  Set potxtFileExp = Nothing
  Set pofsoFileExp = Nothing
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  pgbProgreso.Value = pgbProgreso.Min
  
End Sub
Private Function ppElimina_Informacion() As Boolean
  Dim nRegistro As Long, nRegistros As Long
  Dim sAñoIni As String, sAñoFin As String, sMesIni As String, sMesFin As String
  Dim nSecuencia As Integer, nContador As Integer
  Dim nElemento As Integer, nPicture As Integer
  Dim a_Tabla(), a_Archivo(), a_Where()
  
  ' Cambio el Mensaje y Muestro la Barra
  pgbProgreso.Value = pgbProgreso.Min
  
  sAñoIni = Left(Trim(cmbejercicio(0).Text), 4)
  sAñoFin = Left(Trim(cmbejercicio(1).Text), 4)
  sMesIni = Left(Trim(cmbPeriodo(0).Text), 2)
  sMesFin = Left(Trim(cmbPeriodo(1).Text), 2)
  ' Exporto la información seleccionada
  nContador = outParametro.ListCount - 1
  nRegistros = 1: nElemento = 46
  While nContador > 1
    fMenu.panPercent.Visible = True
    ' Nivel de proceso o selección de información
    If outParametro.Indent(nContador) <> 1 Then
      ' Selecciono el nombre y columnas del archivo de texto
      a_Archivo = Choose(nElemento + 1, Array("cta"), Array("cfg"), Array("cco"), Array("tcb"), _
          Array("bco"), Array("dci"), Array("tpt"), Array("pfs"), Array("dsm"), Array("via"), Array("zon"), _
          Array("cpc"), Array("cls"), Array("afp"), Array("eps"), Array("bas"), Array("cgo"), Array("ubi"), Array("sec"), _
          Array("var"), Array("esq"), Array("pem"), Array("pck"), _
          Array("prc"), Array("cxp"), Array("cpr"), Array("cxc"), Array("cpv"), _
          Array("psn"), Array("rxd"), Array("con"), Array("esr"), Array("exl"), Array("dfm"), _
          Array("boc", "bod"), Array("rpc", "rpd"), Array("rpl", "pld"), _
          Array("pdo"), Array("cte"), Array("asi"), Array("rde"), _
          Array("rhc", "rda"), Array("car"), Array("paf"), _
          Array("vae", "vac", "vad"), Array("gre", "gra"), Array("cse", "csm", "csd", "csc"))
      a_Tabla = Choose(nElemento + 1, Array("cocta"), Array("cocfg"), Array("cocco"), Array("tgtcb"), _
          Array("plbanco"), Array("pldocidentidad"), Array("pltpotrabajador"), Array("plprofesion"), Array("pldstmoneda"), Array("pltipovia"), Array("pltipozona"), _
          Array("plconcepto"), Array("plclasplan"), Array("plentidadafp"), Array("plentidadeps"), Array("pltablabase"), Array("plcargo"), Array("plubicacion"), Array("plseccion"), _
          Array("plvarfunc"), Array("plescalaquinta"), Array("plcfgempresa"), Array("plparametroafp"), _
          Array("plproceso"), Array("plconceplanilla"), Array("plconceproceso"), Array("plctacencos"), Array("plctapvs"), _
          Array("plpersonal"), Array("plremudefa"), Array("plcontrato"), Array("plestudios"), Array("plexpelaboral"), Array("plfamiliares"), _
          Array("plboletapago", "pldetaboleta"), Array("plgenreporte", "pldetareporte"), Array("plplanilla", "pldetaplanilla"), _
          Array("plperiodo"), Array("plcuentacte"), Array("plasistencia"), Array("plremuexce"), _
          Array("plresultado", "pldatoresultado"), Array("plcartabanco"), Array("plplanillafp"), _
          Array("plpvsperiodovac", "plpvsvacacion", "plpvsvacaciondet"), Array("plpvsperiodogra", "plpvsgratifica"), Array("plctsperiodo", "plctsperiodosub", "plctsmovimiento", "plctsresultado"))
      a_Where = Choose(nElemento + 1, Array(""), Array(""), Array(""), Array("WHERE DATE_FORMAT(tcb.fehtcb, '%Y%m')>='" & sAñoIni & sMesIni & "' AND DATE_FORMAT(tcb.fehtcb, '%Y%m')<='" & sAñoFin & sMesFin & "'"), _
          Array(""), Array(""), Array(""), Array(""), Array(""), Array(""), Array(""), _
          Array(""), Array(""), Array(""), Array(""), Array("WHERE bas.pdoano>='" & sAñoIni & "' AND bas.pdoano<='" & sAñoFin & "'"), Array(""), Array(""), Array(""), _
          Array(""), Array(""), Array("WHERE pem.pdoano>='" & sAñoIni & "' AND pem.pdoano>='" & sAñoFin & "'"), Array("WHERE pck.pdoano>='" & sAñoIni & "' AND pck.pdoano>='" & sAñoFin & "'"), _
          Array(""), Array(""), Array(""), Array(""), Array(""), _
          Array("WHERE DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=con.codcls AND psn.codpsn=con.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=esr.codcls AND psn.codpsn=esr.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=exl.codcls AND psn.codpsn=exl.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=dfm.codcls AND psn.codpsn=dfm.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), _
          Array("", ""), Array("", ""), Array("", ""), _
          Array("WHERE CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=cte.codcls AND pdo.codpdo=cte.codpdoprv AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=asi.codcls AND pdo.codpdo=asi.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=rde.codcls AND pdo.codpdo=rde.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE CONCAT(rhc.pdoano, rhc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(rhc.pdoano, rhc.pdomes)<='" & sAñoFin & sMesFin & "'", ", plresultado rhc WHERE rhc.codcls=rda.codcls AND rhc.codpdo=rda.codpdo AND rhc.codpsn=rda.codpsn AND CONCAT(rhc.pdoano, rhc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(rhc.pdoano, rhc.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE DATE_FORMAT(car.fechaproce, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array("WHERE CONCAT(paf.pdoano, paf.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(paf.pdoano, paf.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE vae.codpvs<='" & sAñoFin & "'", "WHERE vac.codpvs<='" & sAñoFin & "'", "WHERE CONCAT(vad.pdoano, vad.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(vad.pdoano, vad.pdomes)<='" & sAñoFin & sMesFin & "'"), Array("WHERE CONCAT(gre.pdoano, gre.mesini)>='" & sAñoIni & sMesIni & "' AND CONCAT(gre.pdoano, gre.mesfin)<='" & sAñoFin & sMesFin & "'", "WHERE CONCAT(gra.pdoano, gra.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(gra.pdoano, gra.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE cse.pdoano>='" & sAñoIni & "' AND cse.pdoano<='" & sAñoFin & "'", ", plctsperiodo cse WHERE cse.codcls=csm.codcls AND cse.pdocts=csm.pdocts AND CONCAT(cse.pdoano, csm.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(cse.pdoano, csm.pdomes)<='" & sAñoFin & sMesFin & "'", ", plctsperiodo cse WHERE cse.codcls=csd.codcls AND cse.pdocts=csd.pdocts AND CONCAT(cse.pdoano, csd.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(cse.pdoano, csd.pdomes)<='" & sAñoFin & sMesFin & "'", "WHERE CONCAT(csc.pdoano, csc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(csc.pdoano, csc.pdomes)<='" & sAñoFin & sMesFin & "'"))
      nPicture = outParametro.PictureType(nContador)
      sfmProgreso.Caption = " Elimina Información: " & Trim(outParametro.List(nContador)) & " "
      
      ' verifico se encuentra seleccionado
      If nPicture = outClosed Then
        nRegistro = 0
        ' Inicializo la barra de progreso
        pgbProgreso.Max = UBound(a_Archivo, 1) + 1
        pgbProgreso.Value = pgbProgreso.Min
        sfmProgreso.Caption = " Elimina Información: " & Trim(outParametro.List(nContador)) & " "
        For nSecuencia = UBound(a_Archivo, 1) To 0 Step -1
          ' Elimino la información existente
          s_Sql = "DELETE " & a_Archivo(nSecuencia) & ".* "
          s_Sql = s_Sql & "FROM " & a_Tabla(nSecuencia) & " " & a_Archivo(nSecuencia) & " "
          s_Sql = s_Sql & IIf(chkParametro.Value, "", a_Where(nSecuencia))
'          Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
          If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finaliza
          ' Incremento el porcentaje
          nRegistro = nRegistro + 1
          pgbProgreso.Value = nRegistro
        Next nSecuencia
      End If
      nElemento = nElemento - 1
    End If
    fMenu.panPercent.FloodPercent = ((nRegistros * 100) \ outParametro.ListCount)
    nContador = nContador - 1
    nRegistros = nRegistros + 1
  Wend
  ppElimina_Informacion = True

Finaliza:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  pgbProgreso.Value = pgbProgreso.Min

End Function
Private Function ppRegistro_Texto(ByVal porstRegistro As ADODB.Recordset, ByVal nColumnas As Integer, bTipo As Byte) As String
  Dim nCampo As Integer, nTipoDato As String, sCaracter As String
  Dim nBinary As Integer
  
  ppRegistro_Texto = "": sCaracter = "|"
  For nCampo = 0 To nColumnas - 1
    nTipoDato = IIf((porstRegistro(nCampo).Type = adSmallInt Or porstRegistro(nCampo).Type = adInteger Or porstRegistro(nCampo).Type = adDouble Or porstRegistro(nCampo).Type = adCurrency Or porstRegistro(nCampo).Type = adNumeric), TipoDato.Numero, IIf(porstRegistro(nCampo).Type = adChar Or porstRegistro(nCampo).Type = adVarChar, TipoDato.Caracter, IIf(porstRegistro(nCampo).Type = adDBDate, TipoDato.FECHA, IIf(porstRegistro(nCampo).Type = adDBTimeStamp, TipoDato.Caracter, TipoDato.Caracter))))
    nBinary = porstRegistro(nCampo).Type
    If bTipo = 2 And (nTipoDato = TipoDato.FECHA Or nBinary = adDBTimeStamp) Then
      ppRegistro_Texto = ppRegistro_Texto & gdl_Funcion.aTexto(IIf(IsNull(porstRegistro(nCampo).Value), "", Format(porstRegistro(nCampo).Value, IIf(nBinary = adDBTimeStamp, s_FmtFeHoMysql_0, s_FormatoFecha)))) & sCaracter
    Else
      ppRegistro_Texto = ppRegistro_Texto & gdl_Funcion.SacaEntRetApos(Choose(bTipo + 1, porstRegistro(nCampo).Name, nTipoDato, IIf((nBinary = adLongVarBinary Or IsNull(porstRegistro(nCampo).Value)), "", porstRegistro(nCampo).Value))) & sCaracter
    End If
  Next nCampo

End Function
Private Function ppRestore_Informacion() As Boolean
  Dim sArchivo As String, psRegistro As String
  Dim nArchivo As Integer, nRegistro As Long, nRegistros As Long
  Dim nSecuencia As Integer, nContador As Integer, nColumna As Integer
  Dim nElemento As Integer, nPicture As Integer
  Dim a_Tabla(), a_Archivo(), a_Columnas(), a_Primary()
  Dim a_Registro(), a_Cabecera(), a_Formato()
  Dim sSQLValor As String, sSeparador As String
  
  ' Cambio el Mensaje y Muestro la Barra
  pgbProgreso.Value = pgbProgreso.Min
  ' Obtengo el archivo de texto libre
  nArchivo = FreeFile
  
  ' Importo la información seleccionada
  nContador = 1: nElemento = 0
  While outParametro.ListCount > nContador
    fMenu.panPercent.Visible = True
    ' Nivel de proceso o selección de información
    If outParametro.Indent(nContador) <> 1 Then
      nPicture = outParametro.PictureType(nContador)
      ' Verifico se encuentra seleccionado
      If nPicture = outClosed Then
        ' Selecciono el nombre y columnas del archivo de texto
        a_Archivo = Choose(nElemento + 1, Array("cta"), Array("cfg"), Array("cco"), Array("tcb"), _
            Array("bco"), Array("dci"), Array("tpt"), Array("pfs"), Array("dsm"), Array("via"), Array("zon"), _
            Array("cpc"), Array("cls"), Array("afp"), Array("eps"), Array("bas"), Array("cgo"), Array("ubi"), Array("sec"), _
            Array("var"), Array("esq"), Array("pem"), Array("pck"), _
            Array("prc"), Array("cxp"), Array("cpr"), Array("cxc"), Array("cpv"), _
            Array("psn"), Array("rxd"), Array("con"), Array("esr"), Array("exl"), Array("dfm"), _
            Array("boc", "bod"), Array("rpc", "rpd"), Array("rpl", "pld"), _
            Array("pdo"), Array("cte"), Array("asi"), Array("rde"), _
            Array("rhc", "rda"), Array("car"), Array("paf"), _
            Array("vae", "vac", "vad"), Array("gre", "gra"), Array("cse", "csm", "csd", "csc"))
        a_Tabla = Choose(nElemento + 1, Array("cocta"), Array("cocfg"), Array("cocco"), Array("tgtcb"), _
            Array("plbanco"), Array("pldocidentidad"), Array("pltpotrabajador"), Array("plprofesion"), Array("pldstmoneda"), Array("pltipovia"), Array("pltipozona"), _
            Array("plconcepto"), Array("plclasplan"), Array("plentidadafp"), Array("plentidadeps"), Array("pltablabase"), Array("plcargo"), Array("plubicacion"), Array("plseccion"), _
            Array("plvarfunc"), Array("plescalaquinta"), Array("plcfgempresa"), Array("plparametroafp"), _
            Array("plproceso"), Array("plconceplanilla"), Array("plconceproceso"), Array("plctacencos"), Array("plctapvs"), _
            Array("plpersonal"), Array("plremudefa"), Array("plcontrato"), Array("plestudios"), Array("plexpelaboral"), Array("plfamiliares"), _
            Array("plboletapago", "pldetaboleta"), Array("plgenreporte", "pldetareporte"), Array("plplanilla", "pldetaplanilla"), _
            Array("plperiodo"), Array("plcuentacte"), Array("plasistencia"), Array("plremuexce"), _
            Array("plresultado", "pldatoresultado"), Array("plcartabanco"), Array("plplanillafp"), _
            Array("plpvsperiodovac", "plpvsvacacion", "plpvsvacaciondet"), Array("plpvsperiodogra", "plpvsgratifica"), Array("plctsperiodo", "plctsperiodosub", "plctsmovimiento", "plctsresultado"))
        a_Primary = Choose(nElemento + 1, Array("codcta"), Array("pdoano"), Array("codcco"), Array("fehtcb"), _
            Array("codbco"), Array("coddci"), Array("codtpt"), Array("codpfs"), Array("codmon"), Array("codvia"), Array("codzona"), _
            Array("codcpc"), Array("codcls"), Array("codafp"), Array("codeps"), Array("codtbl"), Array("codcgo"), Array("codubica"), Array("codsec"), _
            Array("tipo"), Array("orden"), Array("pdoano"), Array("pdoano"), _
            Array("codproce"), Array("codcpc"), Array("codcpc"), Array("codcpc"), Array("codcco"), _
            Array("codpsn"), Array("codpsn"), Array("codpsn"), Array("codpsn"), Array("codpsn"), Array("codpsn"), _
            Array("codboleta", "codboleta"), Array("codrpt", "codrpt"), Array("codpll", "codpll"), _
            Array("codpdo"), Array("numctacte"), Array("codpdo"), Array("codpdo"), _
            Array("codpdo", "codpdo"), Array("nrocarta"), Array("nrohoja"), _
            Array("codpvs", "codpvs", "codpvs"), Array("sempvs", "sempvs"), Array("pdocts", "pdocts", "pdocts", "pdocts"))
        a_Columnas = Choose(nElemento + 1, Array(26), Array(21), Array(8), Array(7), _
            Array(11), Array(8), Array(7), Array(7), Array(8), Array(8), Array(8), _
            Array(9), Array(10), Array(16), Array(9), Array(22), Array(8), Array(7), Array(7), _
            Array(6), Array(7), Array(49), Array(37), _
            Array(8), Array(11), Array(9), Array(14), Array(15), _
            Array(70), Array(9), Array(15), Array(12), Array(12), Array(29), _
            Array(15, 18), Array(12, 14), Array(19, 9), _
            Array(15), Array(25), Array(38), Array(10), _
            Array(22, 18), Array(16), Array(17), _
            Array(10, 12, 26), Array(12, 25), Array(9, 17, 20, 21))
            
        ' Desactivo la opcion seleccionada
        outParametro.PictureType(nContador) = outOpen
        For nSecuencia = 0 To UBound(a_Archivo, 1)
          ' Verifico si existe el archivo de texto y activo la opción
          sArchivo = dlbDirectorio(0).path & "\" & ps_RucEmpresa & a_Archivo(nSecuencia) & ".bma"
          If dir$(sArchivo, vbNormal) <> "" Then
            outParametro.PictureType(nContador) = outClosed
            ' Aperturo el archivo plano
            Open sArchivo For Input As #nArchivo
            nRegistros = CLng(LOF(nArchivo))
            If nRegistros > 0 Then
              ' Redimenciono los arreglos de grabación
              ReDim a_Registro(a_Columnas(nSecuencia))
              ReDim a_Cabecera(a_Columnas(nSecuencia))
              ReDim a_Formato(a_Columnas(nSecuencia))
              
              ' Inicializo la barra de progreso
              pgbProgreso.Max = nRegistros
              pgbProgreso.Value = pgbProgreso.Min
              sfmProgreso.Caption = " Restore Información: " & Trim(outParametro.List(nContador)) & " - " & Right(sArchivo, 18) & " "
              ' Elimino y creo el archivo temporal de grabacion/restauración de información
              s_Sql = "DROP TABLE IF EXISTS tmp" & Mid(a_Tabla(nSecuencia), InStr(a_Tabla(nSecuencia), ".") + 1)
              If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finaliza
              s_Sql = "CREATE TEMPORARY TABLE tmp" & Mid(a_Tabla(nSecuencia), InStr(a_Tabla(nSecuencia), ".") + 1)
              s_Sql = s_Sql & " SELECT *, '999999' AS registro"
              s_Sql = s_Sql & " FROM " & a_Tabla(nSecuencia)
              s_Sql = s_Sql & " WHERE " & a_Primary(nSecuencia) & "='tmpusrma'"
              If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finaliza
              
              ' Selecciono los nombres de los campos o columnas
              Line Input #nArchivo, psRegistro
              Registro_Texto psRegistro, a_Columnas(nSecuencia), a_Cabecera
              nRegistro = 1
              ' Obtengo los tipos de los campos o clumnas
              Line Input #nArchivo, psRegistro
              Registro_Texto psRegistro, a_Columnas(nSecuencia), a_Formato
              nRegistro = 2
              ' Inserto los datos o detalle la tabla temporal
              Do While Not EOF(nArchivo)
                Line Input #nArchivo, psRegistro
                Registro_Texto psRegistro, a_Columnas(nSecuencia), a_Registro
                nRegistro = nRegistro + 1
                
                ' Genero la cadena de grabación
                s_Sql = "INSERT INTO tmp" & Mid(a_Tabla(nSecuencia), InStr(a_Tabla(nSecuencia), ".") + 1) & " ("
                sSQLValor = "VALUES("
                For nColumna = 1 To a_Columnas(nSecuencia)
                  sSeparador = ", "
                  s_Sql = s_Sql & a_Cabecera(nColumna) & sSeparador
                  If a_Formato(nColumna) = TipoDato.Caracter Then
                    a_Registro(nColumna) = gdl_Funcion.SacaEntRetApos(a_Registro(nColumna))
                    a_Registro(nColumna) = Replace(a_Registro(nColumna), "\", "\\")
                    sSQLValor = sSQLValor & IIf(a_Registro(nColumna) = "", "NULL", "'" & a_Registro(nColumna) & "'")
                  ElseIf a_Formato(nColumna) = TipoDato.FECHA Then
                    If IsDate(a_Registro(nColumna)) Then
                      sSQLValor = sSQLValor & "DATE_FORMAT('" & Format(a_Registro(nColumna), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "')"
                    Else
                      sSQLValor = sSQLValor & "NULL"
                    End If
                  ElseIf a_Formato(nColumna) = TipoDato.Numero Then
                    a_Registro(nColumna) = IIf(IsNumeric(a_Registro(nColumna)), a_Registro(nColumna), 0)
                    sSQLValor = sSQLValor & CDec(a_Registro(nColumna))
                  End If
                  sSQLValor = sSQLValor & sSeparador
                Next nColumna
                ' Información del usuario y fecha-hora
                s_Sql = s_Sql & "registro) "
                sSQLValor = sSQLValor & "'" & Format(nRegistro, "000000") & "')"
                s_Sql = s_Sql & sSQLValor
                ' Ejecuto la insercion del registro en la tabla
                If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then Close #nArchivo: GoTo Finaliza
                ' Actualizo la barra de progreso
                pgbProgreso.Value = IIf((Loc(nArchivo) * 128) > nRegistros, nRegistros, (Loc(nArchivo) * 128))
                DoEvents
              Loop
            End If
            ' Cierro el archivo plano
            Close #nArchivo
          End If
        Next nSecuencia
      End If
      ' Incremento detalle de seleccion
      nElemento = nElemento + 1
    End If
    fMenu.panPercent.FloodPercent = ((nContador * 100) \ outParametro.ListCount)
    nContador = nContador + 1
  Wend
  ppRestore_Informacion = True

Finaliza:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  pgbProgreso.Value = pgbProgreso.Min
  
End Function
Private Function ppSincroniza_Informacion() As Boolean
  Dim psRegistro As String, psWhere As String, sArchivo As String
  Dim nRegistro As Long, nRegistros As Long
  Dim sAñoIni As String, sAñoFin As String, sMesIni As String, sMesFin As String
  Dim nSecuencia As Integer, nContador As Integer
  Dim nElemento As Integer, nPicture As Integer
  Dim a_Tabla(), a_Archivo(), a_Columnas(), a_Where(), a_Primary(), a_Orden()
  
  ' Cambio el Mensaje y Muestro la Barra
  pgbProgreso.Value = pgbProgreso.Min
  
  sAñoIni = Left(Trim(cmbejercicio(0).Text), 4)
  sAñoFin = Left(Trim(cmbejercicio(1).Text), 4)
  sMesIni = Left(Trim(cmbPeriodo(0).Text), 2)
  sMesFin = Left(Trim(cmbPeriodo(1).Text), 2)
  ' Exporto la información seleccionada
  nContador = 1: nElemento = 0
  While outParametro.ListCount > nContador
    fMenu.panPercent.Visible = True
    ' Nivel de proceso o selección de información
    If outParametro.Indent(nContador) <> 1 Then
      ' Selecciono el nombre y columnas del archivo de texto
      a_Archivo = Choose(nElemento + 1, Array("cta"), Array("cfg"), Array("cco"), Array("tcb"), _
          Array("bco"), Array("dci"), Array("tpt"), Array("pfs"), Array("dsm"), Array("via"), Array("zon"), _
          Array("cpc"), Array("cls"), Array("afp"), Array("eps"), Array("bas"), Array("cgo"), Array("ubi"), Array("sec"), _
          Array("var"), Array("esq"), Array("pem"), Array("pck"), _
          Array("prc"), Array("cxp"), Array("cpr"), Array("cxc"), Array("cpv"), _
          Array("psn"), Array("rxd"), Array("con"), Array("esr"), Array("exl"), Array("dfm"), _
          Array("boc", "bod"), Array("rpc", "rpd"), Array("rpl", "pld"), _
          Array("pdo"), Array("cte"), Array("asi"), Array("rde"), _
          Array("rhc", "rda"), Array("car"), Array("paf"), _
          Array("vae", "vac", "vad"), Array("gre", "gra"), Array("cse", "csm", "csd", "csc"))
      a_Tabla = Choose(nElemento + 1, Array("cocta"), Array("cocfg"), Array("cocco"), Array("tgtcb"), _
          Array("plbanco"), Array("pldocidentidad"), Array("pltpotrabajador"), Array("plprofesion"), Array("pldstmoneda"), Array("pltipovia"), Array("pltipozona"), _
          Array("plconcepto"), Array("plclasplan"), Array("plentidadafp"), Array("plentidadeps"), Array("pltablabase"), Array("plcargo"), Array("plubicacion"), Array("plseccion"), _
          Array("plvarfunc"), Array("plescalaquinta"), Array("plcfgempresa"), Array("plparametroafp"), _
          Array("plproceso"), Array("plconceplanilla"), Array("plconceproceso"), Array("plctacencos"), Array("plctapvs"), _
          Array("plpersonal"), Array("plremudefa"), Array("plcontrato"), Array("plestudios"), Array("plexpelaboral"), Array("plfamiliares"), _
          Array("plboletapago", "pldetaboleta"), Array("plgenreporte", "pldetareporte"), Array("plplanilla", "pldetaplanilla"), _
          Array("plperiodo"), Array("plcuentacte"), Array("plasistencia"), Array("plremuexce"), _
          Array("plresultado", "pldatoresultado"), Array("plcartabanco"), Array("plplanillafp"), _
          Array("plpvsperiodovac", "plpvsvacacion", "plpvsvacaciondet"), Array("plpvsperiodogra", "plpvsgratifica"), Array("plctsperiodo", "plctsperiodosub", "plctsmovimiento", "plctsresultado"))
      a_Where = Choose(nElemento + 1, Array(""), Array(""), Array(""), Array("WHERE DATE_FORMAT(tcb.fehtcb, '%Y%m')>='" & sAñoIni & sMesIni & "' AND DATE_FORMAT(tcb.fehtcb, '%Y%m')<='" & sAñoFin & sMesFin & "'"), _
          Array(""), Array(""), Array(""), Array(""), Array(""), Array(""), Array(""), _
          Array(""), Array(""), Array(""), Array(""), Array("WHERE bas.pdoano>='" & sAñoIni & "' AND bas.pdoano<='" & sAñoFin & "'"), Array(""), Array(""), Array(""), _
          Array(""), Array(""), Array("WHERE pem.pdoano>='" & sAñoIni & "' AND pem.pdoano>='" & sAñoFin & "'"), Array("WHERE pck.pdoano>='" & sAñoIni & "' AND pck.pdoano>='" & sAñoFin & "'"), _
          Array(""), Array(""), Array(""), Array(""), Array(""), _
          Array("WHERE DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=con.codcls AND psn.codpsn=con.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=esr.codcls AND psn.codpsn=esr.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=exl.codcls AND psn.codpsn=exl.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array(", plpersonal psn WHERE psn.codcls=dfm.codcls AND psn.codpsn=dfm.codpsn AND DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & sAñoFin & sMesFin & "'"), _
          Array("", ""), Array("", ""), Array("", ""), _
          Array("WHERE CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=cte.codcls AND pdo.codpdo=cte.codpdoprv AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=asi.codcls AND pdo.codpdo=asi.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), Array(", plperiodo pdo WHERE pdo.codcls=rde.codcls AND pdo.codpdo=rde.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & sAñoIni & sMesIni & "' AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE CONCAT(rhc.pdoano, rhc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(rhc.pdoano, rhc.pdomes)<='" & sAñoFin & sMesFin & "'", ", plresultado rhc WHERE rhc.codcls=rda.codcls AND rhc.codpdo=rda.codpdo AND rhc.codpsn=rda.codpsn AND CONCAT(rhc.pdoano, rhc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(rhc.pdoano, rhc.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE DATE_FORMAT(car.fechaproce, '%Y%m')<='" & sAñoFin & sMesFin & "'"), Array("WHERE CONCAT(paf.pdoano, paf.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(paf.pdoano, paf.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE vae.codpvs<='" & sAñoFin & "'", "WHERE vac.codpvs<='" & sAñoFin & "'", "WHERE CONCAT(vad.pdoano, vad.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(vad.pdoano, vad.pdomes)<='" & sAñoFin & sMesFin & "'"), Array("WHERE CONCAT(gre.pdoano, gre.mesini)>='" & sAñoIni & sMesIni & "' AND CONCAT(gre.pdoano, gre.mesfin)<='" & sAñoFin & sMesFin & "'", "WHERE CONCAT(gra.pdoano, gra.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(gra.pdoano, gra.pdomes)<='" & sAñoFin & sMesFin & "'"), _
          Array("WHERE cse.pdoano>='" & sAñoIni & "' AND cse.pdoano<='" & sAñoFin & "'", ", plctsperiodo cse WHERE cse.codcls=csm.codcls AND cse.pdocts=csm.pdocts AND CONCAT(cse.pdoano, csm.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(cse.pdoano, csm.pdomes)<='" & sAñoFin & sMesFin & "'", ", plctsperiodo cse WHERE cse.codcls=csd.codcls AND cse.pdocts=csd.pdocts AND CONCAT(cse.pdoano, csd.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(cse.pdoano, csd.pdomes)<='" & sAñoFin & sMesFin & "'", "WHERE CONCAT(csc.pdoano, csc.pdomes)>='" & sAñoIni & sMesIni & "' AND CONCAT(csc.pdoano, csc.pdomes)<='" & sAñoFin & sMesFin & "'"))
      a_Columnas = Choose(nElemento + 1, Array("CodCta, DetCta, DetCtax, TpoCta, NatCta, TpoSdo, TpoAnl, CodCta_Dst_Deb, CodCta_Dst_Hab, TpoMon, TpoTCb, TpoAjD, CodCta_Ajd_Deb, CodCta_Ajd_Hab, IndAjd, CodCta_Crr, IndCCo, IndDoc, IndMoe, IndPsp, Indfjo, EstCta, UsrCre, FyHCre, UsrMdf, FyHMdf"), Array("pdoano, mesatu, tpoMon_fnc, tpoMon_sgn_mn, tpoMon_sgn_me, codCta_nv3, codCta_nv4, codCta_nv5, codCta_nv6, codCta_nv7, codCta_nv8, codtdc_pcp, codtdc_rtc, codcta_pcp, codcta_rtc, indcco, indmne, indrtc, indpcp, codcco_nv3, codcco_nv5"), Array("codcco, detcco, detccox, estcco, usrcre, fyhcre, usrmdf, fyhmdf"), Array("fehtcb, imptcb_Cpr, imptcb_Vta, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codbco, desbco, cuentamn, cuentame, codentidad, formato, estadobco, usrcre, fyhcre, usrmdf, fyhmdf"), Array("coddci, desdci, sigladci, estadodci, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codtpt, destpt, estadotpt, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codpfs, despfs, estadopfs, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codmon, valordmo, desdmo, estadodmo, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codvia, desvia, abrevia, estadovia, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codzona, deszona, abrezona, estadozona, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcpc, descpc, aliascpc, tipocpc, estadocpc, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, descls, clave, horadiaria, fmtboleta, estadocls, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codafp, desafp, factor1, factor2, factor3, factor4, codbco, ctacteafp, desctacteafp, ctactefondo, desctactefondo, estadoafp, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codeps, deseps, ruceps, factoreps, estadoeps, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, pdoano, codtbl, destbl, tpotbl, valordefa, valor01, valor02, valor03, valor04, valor05, valor06, valor07, valor08, valor09, valor10, valor11, valor12, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codcgo, descgo, estadocgo, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codubica, desubica, estadoubica, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codsec, dessec, estadosec, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("tipo, codigo, nombre, descripcion, orden, valor"), Array("orden, numerouit, factor, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("pdoano, codvia, direccionvia, numerodir, codzona, direccionzona, ubigeodir, regpatronal, girocomercial, telefono, email, repapepaterno, repapematerno, repnombres, repcoddci, repnumdocu, psnapepaterno, psnapematerno, psnnombres, psntelefono, codcco, repimpbol, contrato_dot, contrato_doc, rembasica, rempromedio, rempendiente, gratipendiente, remanterior, remganada, codcpc5ta, codtbluit, codtblretener, codtblpendiente, codtbldividir, gratixasis, remxutiejer1, remxutiejer2, remxutiejer3, remxutiejer4, rentaxejer_mn, rentaxejer_me, porcepartici, logo, firma, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("pdoano, cpcbasico, cpcremase, cpcapobli, cpcapovolsfp, cpcapovolcfp, cpcapoemp, cpcseguro, cpcporcen, cpcremuies, cpcremuonp, cpcremuessalud, cpcremuartista, cpcremuquinta, cpcies, cpconp, cpcessalud, cpcartista, cpcquinta, cpceps, remubasicacts, remupromects, remugraticts, remutotalcts, remunepvscts,remudiascts, remumesescts, remuanoscts, remupromeliq, remuvacaliq, remuvacatrun, remuliquiex, remuliquisu, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codproce, desproce, estadoproce, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codcpc, clasecpc, defaultcpc, impbolecpc, formulafun, imagenfun, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codproce, codcpc, secuencia, formulafun, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codcco, codsec, codcpc, orden, codafp, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codcco, orden, codctav_debmn, codctav_habmn, codctav_debme, codctav_habme, codctag_debmn, codctag_habmn, codctag_debme, codctag_habme, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codpsn, apepaterno, apematerno, nombres, fecnacimiento, ubigeonac, nacionalidad, naciextrapsn, sexopsn, refedirec, codvia, nomviadirec, numerdirec, intedirec, codzona, nomzondirec, ubigeodir, estcivilpsn, numhijo, numdepen, coddci, numdociden, numdocmil, telefono, celular, dctojudicial, pordsctojudi, fecingreso, codtpt, codcgo, cgoconfianza, codpfs, codcco, codafp, numeroafp, pagodolar, codbcopago, cuentapago, ctsdeposito, ctsdolar, codbcocts, cuentacts, codeps, regpension, fecingregpen, essvida, cobsctr, afilsindical, remintegralgrati, remintegralvaca, remintegralcts, remimprecisa, remuneta, netocpc, variacpc, imporemuneto, fecbaja, nroessalud, codubica, codsec, coddeudor, codacredor, fecestado, fotopsn, estadopsn, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpsn, codcpc, codmon, imporemune, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpsn, numdocumen, ano, mes, dia, fechaini, fechafin, observacion, archivo, estadocon, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codpsn, orden, institucion, grado, fechaini, fechafin, observacion, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpsn, orden, empresa, codcgo, fechaini, fechafin, observacion, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpsn, orden, apepaterno, apematerno, nombres, fecnacimiento, sexofam, coddci, numdociden, vinculo, cartamed, domicilio, codvia, nomviadom, numerdom, intedom, codzona, nomzonadom, refedom, ubigeodom, incapacidad, certificadomed, motivoina, estadofam, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codboleta, desboleta, orientacion, calidad, papelancho, papelalto, font, copia, lininicopia, estadobol, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, codboleta, seccion, dato, tipodato, fila, columna, longitud, origen, sizefont, fontn, fonts, fontc, desdato, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codrpt, desrpt, formarpt, titulorpt, pierpt, interlinea, anchorpt, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, codrpt, orden, descripcion, tipo, alias, nivel, signo, impreso, longitud, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpll, fila, columna, despll, tipo, alias, descripcion, posicion, longitud, subrayado, sizefont, sizepapel, posipapel, imprimecab, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, codpll, fila, columna, codcpc, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codpdo, despdo, tpopdo, fechaini, fechafin, anopdo, mespdo, estadopdo, fechaproceso, tipocambio, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpsn, numctacte, numcuota, tpoctacte, codcpc, codpdoprv, fectacte, indchecar, numchecar, codbco, indgratifi, tpodscto, codmon, cargo_mn, abono_mn, cargo_me, abono_me, indprn, codpdocan, estadoctacte, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpdo, codpsn, diatrabajo, horanormal, horatipo1, horatipo2, horatipo3, diafalta, tardanza, diaprepostnatal, accidente, diavacaciones, enfermedad, licencia, permisos, fechainivacacion, fechafinvacacion, pdovaca1, fechainivaca1, fechafinvaca1, pdovaca2, fechainivaca2, fechafinvaca2, dialiquidacion, liquidavacacion, diagratificacion, fechacese, fechainiliqvaca, fechafinliqvaca, observacion, tercerturno, suspension, opcional, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codpdo, codpsn, codcpc, codmon, imporemune, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codpdo, codproce, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, codproce_pdo, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, codpdo, codpsn, codcco, codafp, codeps, regpension, naciextrapsn, fecingreso, codubica, codsec, codcgo, fecestado, estadopsn, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, codbco, nrocarta, codcpc, desmotivo, codpsn, codpdo, fechaproce, codmon, importe_mn, importe_me, porinteres, usrcre, fyhcre, usrmdf, fyhmdf"), Array("pdoano, pdomes, codafp, nrohoja, sinpago, fechapago, formapago, interespension, chequepension, codbcopension, interesadmin, chequeadmin, codbcoadmin, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, codpvs, descripvs, fechapvs, rembasbeneficio, estadopvs, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, codpvs, codpsn, pdopvs, fechaini, fechafin, numerodias, estadovac, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, codpvs, codpsn, pdopvs, pdoano, pdomes, fechaini, fechafin, numerodias, codmon, remunera_mn, remunera_me, imporpvsacu_mn, imporpvsacu_me, importepvs_mn, importepvs_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, fechacan, estadodet, usrcre, fyhcre, usrmdf, fyhmdf"), Array("codcls, pdoano, sempvs, descripvs, mesini, mesfin, rembasbeneficio, estadopvs, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, pdoano, sempvs, codpsn, pdomes, fechaini, fechafin, numerodias, codmon, remunera_mn, remunera_me, imporpvsacu_mn, imporpvsacu_me, importepvs_mn, importepvs_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, fechacan, estadogra, usrcre, fyhcre, usrmdf, fyhmdf"), _
          Array("codcls, pdocts, descricts, pdoano, estadocts, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, pdocts, subcts, descrisub, pdomes, numeroanos, numeromeses, numerodias, fechaini, fechafin, fechaven, fechacan, estadosub, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, pdocts, subcts, codpsn, pdomes, numeroanos, numeromeses, numerodias, fechaini, fechafin, fechaven, fechacan, porinteres, estadomov, tipocambio, nrodeposito, usrcre, fyhcre, usrmdf, fyhmdf", "codcls, pdocts, subcts, codpsn, codcpc, secuencia, codmon, importe_mn, importe_me, codcta_debmn, codcta_habmn, codcta_debme, codcta_habme, pdoano, pdomes, tipocpc, impbolecpc, usrcre, fyhcre, usrmdf, fyhmdf"))
      a_Primary = Choose(nElemento + 1, Array("WHERE tmp.codcta=cta.codcta"), Array("WHERE tmp.pdoano=cfg.pdoano"), Array("WHERE tmp.codcco=cco.codcco"), Array("WHERE tmp.fehtcb=tcb.fehtcb"), _
          Array("WHERE tmp.codbco=bco.codbco"), Array("WHERE tmp.coddci=dci.coddci"), Array("WHERE tmp.codtpt=tpt.codtpt"), Array("WHERE tmp.codpfs=pfs.codpfs"), Array("WHERE tmp.codmon=dsm.codmon AND tmp.valordmo=dsm.valordmo"), Array("WHERE tmp.codvia=via.codvia"), Array("WHERE tmp.codzona=zon.codzona"), _
          Array("WHERE tmp.codcpc=cpc.codcpc"), Array("WHERE tmp.codcls=cls.codcls"), Array("WHERE tmp.codafp=afp.codafp"), Array("WHERE tmp.codeps=eps.codeps"), Array("WHERE tmp.codcls=bas.codcls AND tmp.pdoano=bas.pdoano AND tmp.codtbl=bas.codtbl"), Array("WHERE tmp.codcls=cgo.codcls AND tmp.codcgo=cgo.codcgo"), Array("WHERE tmp.codubica=ubi.codubica"), Array("WHERE tmp.codsec=sec.codsec"), _
          Array("WHERE tmp.tipo=var.tipo AND tmp.codigo=var.codigo AND tmp.nombre=var.nombre"), Array("WHERE tmp.orden=esq.orden"), Array("WHERE tmp.pdoano=pem.pdoano"), Array("WHERE tmp.pdoano=pck.pdoano"), _
          Array("WHERE tmp.codcls=prc.codcls AND tmp.codproce=prc.codproce"), Array("WHERE tmp.codcls=cxp.codcls AND tmp.codcpc=cxp.codcpc"), Array("WHERE tmp.codcls=cpr.codcls AND tmp.codproce=cpr.codproce AND tmp.codcpc=cpr.codcpc AND tmp.secuencia=cpr.secuencia"), Array("WHERE tmp.codcls=cxc.codcls AND tmp.codcco=cxc.codcco AND tmp.codsec=cxc.codsec AND tmp.codcpc=cxc.codcpc AND tmp.orden=cxc.orden"), Array("WHERE tmp.codcls=cpv.codcls AND tmp.codcco=cpv.codcco AND tmp.orden=cpv.orden"), _
          Array("WHERE tmp.codcls=psn.codcls AND tmp.codpsn=psn.codpsn"), Array("WHERE tmp.codcls=rxd.codcls AND tmp.codpsn=rxd.codpsn AND tmp.codcpc=rxd.codcpc"), Array("WHERE tmp.codcls=con.codcls AND tmp.codpsn=con.codpsn AND tmp.numdocumen=con.numdocumen AND tmp.ano=con.ano AND tmp.mes=con.mes AND tmp.dia=con.dia"), Array("WHERE tmp.codcls=esr.codcls AND tmp.codpsn=esr.codpsn AND tmp.orden=esr.orden"), Array("WHERE tmp.codcls=exl.codcls AND tmp.codpsn=exl.codpsn AND tmp.orden=exl.orden"), Array("WHERE tmp.codcls=dfm.codcls AND tmp.codpsn=dfm.codpsn AND tmp.orden=dfm.orden"), _
          Array("WHERE tmp.codcls=boc.codcls AND tmp.codboleta=boc.codboleta", "WHERE tmp.codcls=bod.codcls AND tmp.codboleta=bod.codboleta AND tmp.seccion=bod.seccion AND tmp.dato=bod.dato AND tmp.tipodato=bod.tipodato AND tmp.fila=bod.fila AND tmp.columna=bod.columna"), Array("WHERE tmp.codcls=rpc.codcls AND tmp.codrpt=rpc.codrpt", "WHERE tmp.codcls=rpd.codcls AND tmp.codrpt=rpd.codrpt AND tmp.orden=rpd.orden"), Array("WHERE tmp.codcls=rpl.codcls AND tmp.codpll=rpl.codpll AND tmp.fila=rpl.fila AND tmp.columna=rpl.columna", "WHERE tmp.codcls=pld.codcls AND tmp.codpll=pld.codpll AND tmp.fila=pld.fila AND tmp.columna=pld.columna AND tmp.codcpc=pld.codcpc"), _
          Array("WHERE tmp.codcls=pdo.codcls AND tmp.codpdo=pdo.codpdo"), Array("WHERE tmp.codcls=cte.codcls AND tmp.codpsn=cte.codpsn AND tmp.numctacte=cte.numctacte AND tmp.numcuota=cte.numcuota"), Array("WHERE tmp.codcls=asi.codcls AND tmp.codpdo=asi.codpdo AND tmp.codpsn=asi.codpsn"), Array("WHERE tmp.codcls=rde.codcls AND tmp.codpdo=rde.codpdo AND tmp.codpsn=rde.codpsn AND tmp.codcpc=rde.codcpc"), _
          Array("WHERE tmp.codcls=rhc.codcls AND tmp.codpdo=rhc.codpdo AND tmp.codproce=rhc.codproce AND tmp.codpsn=rhc.codpsn AND tmp.codcpc=rhc.codcpc AND tmp.secuencia=rhc.secuencia", "WHERE tmp.codcls=rda.codcls AND tmp.codpdo=rda.codpdo AND tmp.codpsn=rda.codpsn"), Array("WHERE tmp.codcls=car.codcls AND tmp.codbco=car.codbco AND tmp.nrocarta=car.nrocarta AND tmp.codcpc=car.codcpc AND tmp.codpsn=car.codpsn"), Array("WHERE tmp.pdoano=paf.pdoano AND tmp.pdomes=paf.pdomes AND tmp.codafp=paf.codafp"), _
          Array("WHERE tmp.codcls=vae.codcls AND tmp.codpvs=vae.codpvs", "WHERE tmp.codcls=vac.codcls AND tmp.codpvs=vac.codpvs AND tmp.codpsn=vac.codpsn AND tmp.pdopvs=vac.pdopvs", "WHERE tmp.codcls=vad.codcls AND tmp.codpvs=vad.codpvs AND tmp.codpsn=vad.codpsn AND tmp.pdopvs=vad.pdopvs AND tmp.pdoano=vad.pdoano AND tmp.pdomes=vad.pdomes"), Array("WHERE tmp.codcls=gre.codcls AND tmp.pdoano=gre.pdoano AND tmp.sempvs=gre.sempvs", "WHERE tmp.codcls=gra.codcls AND tmp.pdoano=gra.pdoano AND tmp.sempvs=gra.sempvs AND tmp.codpsn=gra.codpsn AND tmp.pdomes=gra.pdomes"), Array("WHERE tmp.codcls=cse.codcls AND tmp.pdocts=cse.pdocts", "WHERE tmp.codcls=csm.codcls AND tmp.pdocts=csm.pdocts AND tmp.subcts=csm.subcts", "WHERE tmp.codcls=csd.codcls AND tmp.pdocts=csd.pdocts AND tmp.subcts=csd.subcts AND tmp.codpsn=csd.codpsn", "WHERE tmp.codcls=csc.codcls AND tmp.pdocts=csc.pdocts AND tmp.subcts=csc.subcts AND tmp.codpsn=csc.codpsn AND tmp.codcpc=csc.codcpc AND tmp.secuencia=csc.secuencia"))
      a_Orden = Choose(nElemento + 1, Array("codcta"), Array("pdoano"), Array("codcco"), Array("fehtcb"), _
          Array("codbco"), Array("coddci"), Array("codtpt"), Array("codpfs"), Array("codmon, valordmo"), Array("codvia"), Array("codzona"), _
          Array("codcpc"), Array("codcls"), Array("codafp"), Array("codeps"), Array("codcls, pdoano, codtbl"), Array("codcls, codcgo"), Array("codubica"), Array("codsec"), _
          Array("tipo, codigo, nombre"), Array("orden"), Array("pdoano"), Array("pdoano"), _
          Array("codcls, codproce"), Array("codcls, codcpc"), Array("codcls, codproce, codcpc, secuencia"), Array("codcls, codcco, codsec, codcpc, orden"), Array("codcls, codcco, orden"), _
          Array("codcls, codpsn"), Array("codcls, codpsn, codcpc"), Array("codcls, codpsn, numdocumen, ano, mes, dia"), Array("codcls, codpsn, orden"), Array("codcls, codpsn, orden"), Array("codcls, codpsn, orden"), _
          Array("codcls, codboleta", "codcls, codboleta, seccion, dato, tipodato, fila, columna"), Array("codcls, codrpt", "codcls, codrpt, orden"), Array("codcls, codpll, fila, columna", "codcls, codpll, fila, columna, codcpc"), _
          Array("codcls, codpdo"), Array("codcls, codpsn, numctacte, numcuota"), Array("codcls, codpdo, codpsn"), Array("codcls, codpdo, codpsn, codcpc"), _
          Array("codcls, codpdo, codproce, codpsn, codcpc, secuencia", "codcls, codpdo, codpsn"), Array("codcls, codbco, nrocarta, codcpc, codpsn"), Array("pdoano, pdomes, codafp"), _
          Array("codcls, codpvs", "codcls, codpvs, codpsn, pdopvs", "codcls, codpvs, codpsn, pdopvs, pdoano, pdomes"), Array("codcls, pdoano, sempvs", "codcls, pdoano, sempvs, codpsn, pdomes"), Array("codcls, pdocts", "codcls, pdocts, subcts", "codcls, pdocts, subcts, codpsn", "codcls, pdocts, subcts, codpsn, codcpc, secuencia"))
      nPicture = outParametro.PictureType(nContador)
      ' verifico se encuentra seleccionado
      If nPicture = outClosed Then
        ' Inicializo la barra de progreso
        nRegistros = UBound(a_Archivo, 1) + 1: nRegistro = 0
        pgbProgreso.Max = nRegistros
        pgbProgreso.Value = pgbProgreso.Min
        sfmProgreso.Caption = " Sincroniza Información: " & Trim(outParametro.List(nContador)) & " "
        For nSecuencia = 0 To UBound(a_Archivo, 1)
          sArchivo = dlbDirectorio(0).path & "\" & ps_RucEmpresa & a_Archivo(nSecuencia) & ".bma"
          If dir$(sArchivo, vbNormal) <> "" Then
            psRegistro = Replace(a_Columnas(nSecuencia), ", ", ", " & a_Archivo(nSecuencia) & ".")
            psWhere = IIf(chkParametro.Value, "", a_Where(nSecuencia) & " ")
            ' Inserto la información no existente a la tabla
            s_Sql = "INSERT INTO " & a_Tabla(nSecuencia) & " (" & a_Columnas(nSecuencia) & ") "
            s_Sql = s_Sql & "SELECT DISTINCTROW " & a_Archivo(nSecuencia) & "." & psRegistro & " "
            s_Sql = s_Sql & "FROM tmp" & a_Tabla(nSecuencia) & " " & a_Archivo(nSecuencia) & " "
            s_Sql = s_Sql & psWhere
            s_Sql = s_Sql & IIf(Trim(psWhere) = "", "WHERE ", "AND ")
            s_Sql = s_Sql & "NOT EXISTS(SELECT tmp.* "
            s_Sql = s_Sql & "FROM " & a_Tabla(nSecuencia) & " tmp "
            s_Sql = s_Sql & a_Primary(nSecuencia) & ") "
            s_Sql = s_Sql & "ORDER BY " & a_Orden(nSecuencia)
            If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finaliza
          End If
          ' Incremento el porcentaje
          nRegistro = nRegistro + 1
          pgbProgreso.Value = nRegistro
        Next nSecuencia
      End If
      nElemento = nElemento + 1
    End If
    fMenu.panPercent.FloodPercent = ((nContador * 100) \ outParametro.ListCount)
    nContador = nContador + 1
  Wend
  ppSincroniza_Informacion = True

Finaliza:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  pgbProgreso.Value = pgbProgreso.Min

End Function
']
Private Sub chkParametro_Click(Value As Integer)
  If chkParametro.Value Then
    For n_Index = 1 To outParametro.ListCount - 1: outParametro.PictureType(n_Index) = outClosed: Next n_Index
  End If
  frmCuadro(0).Enabled = (chkParametro.Value = vbUnchecked)
  cmbejercicio(0).Enabled = (chkParametro.Value = vbUnchecked)
  cmbejercicio(1).Enabled = (chkParametro.Value = vbUnchecked)
  cmbPeriodo(0).Enabled = (chkParametro.Value = vbUnchecked)
  cmbPeriodo(1).Enabled = (chkParametro.Value = vbUnchecked)

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_OldMessage As String
  
  If Index = 1 Then Unload Me: Exit Sub
  ' Verifico que existan registros
  If cmbPeriodo(0).Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo Inicial", vbExclamation: cmbPeriodo(0).SetFocus: Exit Sub
  If cmbPeriodo(1).Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo Final", vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
  If cmbPeriodo(0).ListIndex > cmbPeriodo(1).ListIndex Then Beep: MsgBox "El Periodo de Inicial debe ser menor o igual al Periodo Final", vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de " & IIf(optParametro(0).Value, "Generar archivo de Backup de información ?", " Eliminar información existente de las tablas seleccionadas y las Transacciones por periodos; Restaurar la información del archivo de Backup ?"), vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    pgbProgreso.Min = 0
    Me.Height = 6560
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    
    ' Inicializo Sistema de Contabilidad
    s_Sql = "SELECT DISTINCTROW emp.codemp, emp.siscon "
    s_Sql = s_Sql & "FROM tgemp emp "
    s_Sql = s_Sql & "WHERE emp.codemp='" & ps_CodEmpresa & "' "
    s_Sql = s_Sql & "AND emp.sispla='" & s_Estado_Act & "' "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_BDSystems, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If porstRecordset!siscon = s_Estado_Act Then
      outParametro.PictureType(1) = outOpen
      n_Index = 2
      Do While outParametro.Indent(n_Index) <> 1
        outParametro.PictureType(n_Index) = outOpen
        n_Index = n_Index + 1
        If n_Index = outParametro.ListCount Then Exit Do
      Loop
    End If
    
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    ' Proceso la informacion
    If optParametro(0).Value Then
      ppBackup_Informacion
    Else
      ' Realizo la transformación de información
      If Not ppRestore_Informacion Then GoTo Error
      ' Realizo la verificación de inicialización
      If Not ppElimina_Informacion Then GoTo Error
      ' Realizo la eliminacion(validacion) y restauracion
      If Not ppSincroniza_Informacion Then GoTo Error
    End If
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    MsgBox "Proceso de" & IIf(optParametro(0).Value, " Backup de información ", " Restauración de información del archivo de Backup ") & "; Finalizo con Exito", vbInformation
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  Me.Height = 6010
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub drbUnidad_Change(Index As Integer)
  dlbDirectorio(Index).path = drbUnidad(Index).drive
  dlbDirectorio(Index).Refresh
End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim s_Archivo As String

  'Establece posición y titulo del formulario
  Me.Height = 6020: Me.Width = 7410
  Me.Left = 1080: Me.Top = 200
  
  ' Titulo del formulario y panel
  s_TitleWindow = Me.Caption
  lblTitle = "Información del sistema"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(2, 1) = "proceso": aElemento(2, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 1
    aElemento(n_Index, 1) = Choose(n_Index + 1, "link", "cancelar")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Procesar ", "Cancelar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(1).Cancel = True
  
  drbUnidad(0).drive = ps_PathSystem
  dlbDirectorio(0).path = ps_PathSystem
 
  ' Configuro los graficos de objeto
  s_Archivo = gdl_Procedure.ps_PathImagen & "database.bmp"
  If dir$(s_Archivo, vbNormal) <> "" Then
    imgGrafico.Picture = LoadPicture(s_Archivo)
  End If
  imgGrafico.Refresh
  outParametro.PictureLeaf = imgGrafico
  
  s_Archivo = gdl_Procedure.ps_PathImagen & "nocheck.bmp"
  If dir$(s_Archivo, vbNormal) <> "" Then
    imgGrafico.Picture = LoadPicture(s_Archivo)
  End If
  imgGrafico.Refresh
  outParametro.PictureOpen = imgGrafico
    
  s_Archivo = gdl_Procedure.ps_PathImagen & "check.bmp"
  If dir$(s_Archivo, vbNormal) <> "" Then
    imgGrafico.Picture = LoadPicture(s_Archivo)
  End If
  imgGrafico.Refresh
  outParametro.PictureClosed = imgGrafico
  
  s_Archivo = gdl_Procedure.ps_PathImagen & "mencheck.bmp"
  If dir$(s_Archivo, vbNormal) <> "" Then
    imgGrafico.Picture = LoadPicture(s_Archivo)
  End If
  imgGrafico.Refresh
  outParametro.PictureMinus = imgGrafico
  
  ' Configuro el objeto de parametro
  outParametro.Style = outPlusPictureText
  outParametro.AddItem " Contenido Backup/Restore"
  outParametro.Indent(0) = 0
  outParametro.PictureType(0) = outLeaf
  
  ' Configuración de contabilidad
  outParametro.ListIndex = -1
  outParametro.AddItem " Parametros de Contabilidad"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Plan de Cuentas"
  outParametro.AddItem " Configuración Contabilidad"
  outParametro.AddItem " Centro de Costos"
  outParametro.AddItem " Tipos de Cambio"
  
  ' Tablas generales
  outParametro.ListIndex = -1
  outParametro.AddItem " Tablas Generales"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Entidad Bancarias"
  outParametro.AddItem " Documentos de Identidad"
  outParametro.AddItem " Tipo de Trabajador"
  outParametro.AddItem " Profesión u Oficio"
  outParametro.AddItem " Distribución Monetaria"
  outParametro.AddItem " Tipo de Via - Dirección"
  outParametro.AddItem " Tipo de Zona - Dirección"

  ' Tablas clase planilla
  outParametro.ListIndex = -1
  outParametro.AddItem " Tablas Clase Planilla"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Conceptos de Cálculo"
  outParametro.AddItem " Clase Planilla"
  outParametro.AddItem " Entidad de Pensión"
  outParametro.AddItem " Entidad Prestadora de Salud"
  outParametro.AddItem " Tablas Básicas"
  outParametro.AddItem " Tipo de Cargo de Trabajo"
  outParametro.AddItem " Ubicación o Localidad"
  outParametro.AddItem " Sección Empresarial"

  ' Parametros generales
  outParametro.ListIndex = -1
  outParametro.AddItem " Factores y Constantes"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Parametros y Funciones"
  outParametro.AddItem " Escala de Quinta"
  outParametro.AddItem " Parametros de Empresa"
  outParametro.AddItem " Parametros de Certificados"

  ' Conceptos y procesos de calculo
  outParametro.ListIndex = -1
  outParametro.AddItem " Procesos y Conceptos"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Proceso de Cálculo"
  outParametro.AddItem " Conceptos por Clase Planilla"
  outParametro.AddItem " Conceptos por Proceso de Cálculo"
  outParametro.AddItem " Distribución de Cuentas Contables - Procesos"
  outParametro.AddItem " Distribución de Cuentas Contables - Provisiones"

  ' Datos del personal
  outParametro.ListIndex = -1
  outParametro.AddItem " Datos Personales"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Padrón de Personal"
  outParametro.AddItem " Remuneración/Descuento Default"
  outParametro.AddItem " Contratos de Trabajo"
  outParametro.AddItem " Estudios Realizados"
  outParametro.AddItem " Experiencia Laboral"
  outParametro.AddItem " Datos Familiares"
  
  ' Datos del personal
  outParametro.ListIndex = -1
  outParametro.AddItem " Configuración de Reportes"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Formato boletas"
  outParametro.AddItem " Formato de reportes"
  outParametro.AddItem " Formato de planilla general"

  ' Datos del personal
  outParametro.ListIndex = -1
  outParametro.AddItem " Parametros de Cálculo"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Periodos de Pago"
  outParametro.AddItem " Cuenta Corriente - Prestamos"
  outParametro.AddItem " Control de Asistencia"
  outParametro.AddItem " Remuneración/Descuento Exepcionales"
  
  ' Resultado de calculo
  outParametro.ListIndex = -1
  outParametro.AddItem " Resultado de Cálculo"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Historico de Cálculo - planilla"
  outParametro.AddItem " Cartas de Transferencia - bancos"
  outParametro.AddItem " Planilla afp"
        
  ' Provisiones
  outParametro.ListIndex = -1
  outParametro.AddItem " Proceso de Provisiones"
  outParametro.ListIndex = outParametro.ListCount - 1
  outParametro.AddItem " Provisión de Vacaciones"
  outParametro.AddItem " Provisión de Gratificación"
  outParametro.AddItem " Provisión de CTS"
  ']
  
  ' Configuro los listados, datos adicionales
  For n_Index = (Val(ps_Anyo) - 2) To (Val(ps_Anyo) + 2)
    cmbejercicio(0).AddItem Format(n_Index, "0000")
    cmbejercicio(1).AddItem Format(n_Index, "0000")
  Next n_Index
  cmbejercicio(0).ListIndex = 2: cmbejercicio(1).ListIndex = 2
  For n_Index = 1 To 12
    cmbPeriodo(0).AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre")
    cmbPeriodo(1).AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre")
  Next n_Index
  cmbPeriodo(0).ListIndex = Month(Date) - 1
  cmbPeriodo(1).ListIndex = Month(Date) - 1
  optParametro(0).Value = vbChecked
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub
Private Sub optParametro_Click(Index As Integer, Value As Integer)
  frmCuadro(0).Enabled = (Index = 0)
  cmbPeriodo(0).Enabled = (Index = 0)
  cmbPeriodo(1).Enabled = (Index = 0)
  chkParametro_Click True
End Sub
Private Sub outParametro_PictureClick(ListIndex As Integer)
  Dim nPicture As Integer
  ' Actualizo la selección de parametro
  If outParametro.Indent(ListIndex) > 1 Then
    nPicture = outParametro.PictureType(ListIndex)
    outParametro.PictureType(ListIndex) = Choose(nPicture + 1, outOpen, outClosed)
  End If
End Sub
Private Sub outParametro_PictureDblClick(ListIndex As Integer)
  Dim nPicture As Integer
  ' Actualizo la selección y niveles
  If outParametro.Indent(ListIndex) = 1 Then
    outParametro.ListIndex = ListIndex
    nPicture = outParametro.PictureType(ListIndex)
    outParametro.PictureType(ListIndex) = Choose(nPicture + 1, outOpen, outClosed)
    n_Index = ListIndex + 1
    Do While outParametro.Indent(n_Index) <> 1
      outParametro.PictureType(n_Index) = Choose(nPicture + 1, outOpen, outClosed)
      n_Index = n_Index + 1
      If n_Index = outParametro.ListCount Then Exit Do
    Loop
  End If
End Sub
