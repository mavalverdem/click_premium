VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPTraInf 
   Caption         =   "[título]"
   ClientHeight    =   6720
   ClientLeft      =   1815
   ClientTop       =   1440
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5745
   Begin TabDlg.SSTab tabProceso 
      Height          =   4860
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8573
      _Version        =   393216
      Style           =   1
      TabHeight       =   494
      TabCaption(0)   =   "Importación"
      TabPicture(0)   =   "frmPTraInf.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmTablas(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmProceso(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmUbicacion(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Transferencia"
      TabPicture(1)   =   "frmPTraInf.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmTablas(1)"
      Tab(1).Control(1)=   "frmUbicacion(1)"
      Tab(1).Control(2)=   "frmTablas(3)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Centralización"
      TabPicture(2)   =   "frmPTraInf.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCentra"
      Tab(2).Control(1)=   "frmTablas(2)"
      Tab(2).Control(2)=   "frmProceso(2)"
      Tab(2).Control(3)=   "frmUbicacion(2)"
      Tab(2).ControlCount=   4
      Begin VB.Frame frmTablas 
         Caption         =   "Transferir Información  "
         ForeColor       =   &H00C00000&
         Height          =   1515
         Index           =   3
         Left            =   -74880
         TabIndex        =   74
         Top             =   3240
         Width           =   5295
         Begin VB.ComboBox cboEjercicio 
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
            Height          =   315
            ItemData        =   "frmPTraInf.frx":0054
            Left            =   105
            List            =   "frmPTraInf.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   1110
            Width           =   1755
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   4995
            Picture         =   "frmPTraInf.frx":0058
            Style           =   1  'Graphical
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   510
            Width           =   255
         End
         Begin VB.TextBox txtDato 
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
            Left            =   105
            TabIndex        =   76
            Top             =   510
            Width           =   585
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Empresa :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   80
            Top             =   240
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Ejercicio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   79
            Top             =   885
            Width           =   690
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
            Height          =   315
            Index           =   0
            Left            =   675
            TabIndex        =   78
            Top             =   510
            Width           =   4320
         End
      End
      Begin VB.Frame fraCentra 
         Caption         =   " Parámetros "
         Height          =   1485
         Left            =   -72105
         TabIndex        =   65
         Top             =   2700
         Width           =   2415
         Begin VB.ComboBox cmbParametro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   315
            Index           =   1
            Left            =   1095
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   820
            Width           =   1100
         End
         Begin VB.ComboBox cmbParametro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   315
            Index           =   0
            Left            =   1095
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   390
            Width           =   1100
         End
         Begin VB.Label lblTexto 
            Alignment       =   1  'Right Justify
            Caption         =   "Sucursal :"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   69
            Top             =   850
            Width           =   800
         End
         Begin VB.Label lblTexto 
            Alignment       =   1  'Right Justify
            Caption         =   "Compañia :"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   68
            Top             =   420
            Width           =   800
         End
      End
      Begin VB.Frame frmTablas 
         Caption         =   " Tablas "
         Height          =   2000
         Index           =   2
         Left            =   -74880
         TabIndex        =   58
         Top             =   350
         Width           =   2655
         Begin VB.CommandButton cmdTablas 
            BackColor       =   &H80000002&
            Caption         =   "..."
            Height          =   240
            Index           =   2
            Left            =   2320
            TabIndex        =   72
            ToolTipText     =   "Actualiza Datos de Centro de Costo"
            Top             =   1310
            Width           =   240
         End
         Begin VB.CommandButton cmdTablas 
            BackColor       =   &H80000002&
            Caption         =   "..."
            Height          =   240
            Index           =   1
            Left            =   2320
            TabIndex        =   71
            ToolTipText     =   "Actualiza Datos de Tipo de Documento"
            Top             =   1050
            Width           =   240
         End
         Begin VB.CommandButton cmdTablas 
            BackColor       =   &H80000002&
            Caption         =   "..."
            Height          =   240
            Index           =   0
            Left            =   2320
            TabIndex        =   70
            ToolTipText     =   "Actualiza Datos de Diario"
            Top             =   270
            Width           =   240
         End
         Begin VB.CheckBox chkCentraTabla 
            Caption         =   "&Diario"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   64
            Top             =   270
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraTabla 
            Caption         =   "&Plan de Cuentas"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   63
            Top             =   530
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraTabla 
            Caption         =   "&Auxiliares"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   62
            Top             =   790
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraTabla 
            Caption         =   "&Tipo de Documentos"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   61
            Top             =   1050
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraTabla 
            Caption         =   "&Centro de Costo"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   60
            Top             =   1310
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraTabla 
            Caption         =   "Tipo de Ca&mbio"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   59
            Top             =   1570
            Value           =   1  'Checked
            Width           =   2000
         End
      End
      Begin VB.Frame frmProceso 
         Caption         =   " Transacciones "
         Height          =   1780
         Index           =   2
         Left            =   -74880
         TabIndex        =   53
         Top             =   2400
         Width           =   2655
         Begin VB.CheckBox chkCentraProceso 
            Caption         =   "&Registro de Diario"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   57
            Top             =   1095
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraProceso 
            Caption         =   "Registro de &Honorarios"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   56
            Top             =   840
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraProceso 
            Caption         =   "Registro de &Ventas"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   55
            Top             =   570
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkCentraProceso 
            Caption         =   "Registro de &Compras"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   54
            Top             =   315
            Value           =   1  'Checked
            Width           =   2000
         End
      End
      Begin VB.Frame frmUbicacion 
         Caption         =   " Carpeta "
         Height          =   2280
         Index           =   2
         Left            =   -72120
         TabIndex        =   49
         Top             =   350
         Width           =   2535
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Index           =   2
            Left            =   150
            TabIndex        =   51
            Top             =   690
            Width           =   2235
         End
         Begin VB.DriveListBox drvUnidad 
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   50
            Top             =   400
            Width           =   2235
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   52
            Top             =   200
            Width           =   765
         End
      End
      Begin VB.Frame frmUbicacion 
         Caption         =   " Carpeta "
         Height          =   2880
         Index           =   1
         Left            =   -72120
         TabIndex        =   45
         Top             =   350
         Width           =   2535
         Begin VB.DirListBox dlbDirectorio 
            Height          =   2115
            Index           =   1
            Left            =   150
            TabIndex        =   47
            Top             =   690
            Width           =   2235
         End
         Begin VB.DriveListBox drvUnidad 
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   46
            Top             =   400
            Width           =   2235
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   48
            Top             =   200
            Width           =   765
         End
      End
      Begin VB.Frame frmTablas 
         Caption         =   " Tablas "
         Height          =   2880
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Top             =   350
         Width           =   2655
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Otras Tablas"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   44
            Top             =   2580
            Value           =   1  'Checked
            Width           =   2340
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "Estados &Financiero"
            Height          =   240
            Index           =   7
            Left            =   150
            TabIndex        =   42
            Top             =   2070
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "A&siento Tipo"
            Height          =   240
            Index           =   8
            Left            =   150
            TabIndex        =   43
            Top             =   2325
            Value           =   1  'Checked
            Width           =   2340
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "Tipo de Ca&mbio"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   41
            Top             =   1830
            Value           =   1  'Checked
            Width           =   2340
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Entidad Bancaria"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   36
            Top             =   530
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Plan de Cuentas"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   37
            Top             =   790
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Tipo de Documentos"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   38
            Top             =   1050
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Auxiliares"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   39
            Top             =   1310
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Centro de Costo"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   40
            Top             =   1570
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkTransTabla 
            Caption         =   "&Diario"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   35
            Top             =   270
            Value           =   1  'Checked
            Width           =   2000
         End
      End
      Begin VB.Frame frmUbicacion 
         Caption         =   " Carpeta "
         Height          =   4275
         Index           =   0
         Left            =   2880
         TabIndex        =   21
         Top             =   350
         Width           =   2535
         Begin VB.DriveListBox drvUnidad 
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   33
            Top             =   400
            Width           =   2235
         End
         Begin VB.FileListBox flbArchivo 
            Height          =   1845
            Left            =   150
            Pattern         =   "*.txt"
            TabIndex        =   23
            Top             =   2355
            Width           =   2235
         End
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Index           =   0
            Left            =   150
            TabIndex        =   22
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Archivos :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   32
            Top             =   2150
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   31
            Top             =   200
            Width           =   765
         End
      End
      Begin VB.Frame frmProceso 
         Caption         =   " Transacciones "
         Height          =   2205
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   2655
         Begin VB.CheckBox chkPeriodos 
            Alignment       =   1  'Right Justify
            Caption         =   "Multiples periodos"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   150
            TabIndex        =   75
            Top             =   1365
            Width           =   2055
         End
         Begin VB.CheckBox chkVerificar 
            Alignment       =   1  'Right Justify
            Caption         =   "Verificar Equivalente"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   150
            TabIndex        =   73
            Top             =   1875
            Width           =   2055
         End
         Begin VB.CheckBox chkEliminar 
            Caption         =   "Check1"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   3
            Left            =   2340
            TabIndex        =   19
            Top             =   1095
            Width           =   165
         End
         Begin VB.CheckBox chkEliminar 
            Caption         =   "Check1"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   2
            Left            =   2340
            TabIndex        =   17
            Top             =   840
            Width           =   165
         End
         Begin VB.CheckBox chkEliminar 
            Caption         =   "Check1"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   1
            Left            =   2340
            TabIndex        =   15
            Top             =   570
            Width           =   165
         End
         Begin VB.CheckBox chkEliminar 
            Caption         =   "Check1"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   2340
            TabIndex        =   13
            Top             =   315
            Width           =   165
         End
         Begin VB.CheckBox chkCorrelativo 
            Alignment       =   1  'Right Justify
            Caption         =   "Enumerar Comprobantes"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   150
            TabIndex        =   20
            Top             =   1620
            Width           =   2055
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Registro de &Compras"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   12
            Top             =   315
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Registro de &Ventas"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   14
            Top             =   570
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "Registro de &Honorarios"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   16
            Top             =   840
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporProceso 
            Caption         =   "&Registro de Diario"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   18
            Top             =   1095
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.Label lblEliminar 
            Caption         =   "Eliminar"
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   1980
            TabIndex        =   11
            Top             =   105
            Width           =   570
         End
      End
      Begin VB.Frame frmTablas 
         Caption         =   " Tablas "
         Height          =   2000
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   350
         Width           =   2655
         Begin VB.CheckBox chkModificar 
            Caption         =   "Check1"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   1
            Left            =   2340
            TabIndex        =   4
            Top             =   530
            Width           =   165
         End
         Begin VB.CheckBox chkModificar 
            Caption         =   "Check1"
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   2
            Left            =   2340
            TabIndex        =   6
            Top             =   790
            Width           =   165
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "Tipo de Ca&mbio"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   9
            Top             =   1570
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Centro de Costo"
            Height          =   240
            Index           =   4
            Left            =   150
            TabIndex        =   8
            Top             =   1310
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Tipo de Documentos"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   7
            Top             =   1050
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Auxiliares"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   790
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Plan de Cuentas"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   530
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.CheckBox chkImporTabla 
            Caption         =   "&Diario"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   270
            Value           =   1  'Checked
            Width           =   2000
         End
         Begin VB.Label lblModificar 
            Caption         =   "Modificar"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   1860
            TabIndex        =   1
            Top             =   105
            Width           =   675
         End
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   400
      Left            =   1380
      TabIndex        =   30
      Top             =   6225
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   400
      Left            =   3060
      TabIndex        =   29
      Top             =   6255
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   5880
      Width           =   5535
      _ExtentX        =   9763
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
      TabIndex        =   27
      Top             =   5640
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
      TabIndex        =   25
      Top             =   5040
      Width           =   2310
   End
End
Attribute VB_Name = "frmPTraInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Private aValidar() As Integer
'[Propio del formulario.
Private porstEmpresa As ADODB.Recordset

']
Private Sub cmdAceptar_Click()
'  On Error GoTo Err
    
    'ini 2015-07-09 control flag mayoriza
    Dim xxchkImporProceso As Boolean
    xxchkImporProceso = False
    Dim i As Integer
    For i = 0 To chkImporProceso().UBound
        If chkImporProceso(i).Value Then
            xxchkImporProceso = True
            Exit For
        End If
    Next
    If xxchkImporProceso Then
        'ini 2015-07-27 error trans solo un meses If gcCierre(gsAnoAct, gsMesAct) = 1 Then Exit Sub
        If (gbCieCpb And xxchkImporProceso) Then MsgBox TEXT_9016, vbCritical: cmdSalir.SetFocus: Exit Sub
    End If
    'fin 2015-07-09 control flag mayoriza

    
  Dim porstMRp As New ADODB.Recordset
  Dim nValidacion As Integer
  Dim s_Conexion As String, sSentencia As String

  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  
  pgbProgreso(0).Value = 0: pgbProgreso(0).Min = 0
  pgbProgreso(1).Value = 0: pgbProgreso(1).Min = 0
    
  ' Seteo y activo la coneccion
  s_Conexion = CONNSTRG & gsNomBDS
  If tabProceso.Tab = 2 Then s_Conexion = gfParaOracle
  ' Seteo y activo la coneccion
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    If .State = adStateOpen Then .Close
    .ConnectionTimeout = 15
    .CursorLocation = adUseClient
    .ConnectionString = s_Conexion
    .Open
  End With
                        
  Select Case tabProceso.Tab
   Case 0
    ' Validación de periodos y eliminar
    If (chkPeriodos.Value = vbChecked And chkEliminar(3).Value = vbChecked) Then
      MsgBox Choose(gsIdioma, "Desactive opción de eliminar información comprobantes", "Disable option eliminate information vouchers"), vbInformation, Me.Caption
      cmdAceptar.Enabled = True
      cmdSalir.Enabled = True
      Exit Sub
    End If
    
    ' Validación de proceso
    If (chkEliminar(0).Value = vbChecked Or chkEliminar(1).Value = vbChecked Or chkEliminar(2).Value = vbChecked Or chkEliminar(3).Value = vbChecked) Then
      frmSeguridad.Show vbModal
      If lblEliminar.Tag = ESTCTA_INA Then
        MsgBox Choose(gsIdioma, "Desactive opción de eliminar información", "Disable option eliminate information"), vbInformation, Me.Caption
        cmdAceptar.Enabled = True
        cmdSalir.Enabled = True
        Exit Sub
      End If
    End If
    
    sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#trptRPTraInf_') DROP TABLE #trptRPTraInf"
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRPTraInf", sSentencia)
    sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptRPTraInf (", "CREATE TABLE #trptRPTraInf (")
    sSentencia = sSentencia & "opcion char(1) Null, desopcion varchar(40) Null, caso char(2) Null, "
    sSentencia = sSentencia & "descripcion varchar(80) Null, registro varchar(6) DEFAULT '0')"
    pocnnMain.Execute sSentencia
    
    pocnnMain.BeginTrans               'INICIA TRANSACCION.
    ' Paso 1: Realizo la importación de las tablas y validacione de las mismas generales
    pgbProgreso(0).Max = 6
    pgbProgreso(0).Value = pgbProgreso(0).Min
    ppImporta_Tablas
    pgbProgreso(0).Value = 1                        ' Actualizo la barra de progreso
    
    nValidacion = pfValida_Tablas
    pgbProgreso(0).Value = 2                        ' Actualizo la barra de progreso
    ' Paso 2 : Realizo la importación de las tablas de transacciones
    ppImporta_Proceso
    pgbProgreso(0).Value = 3                        ' Actualizo la barra de progreso
    nValidacion = pfValida_Proceso(nValidacion)
    pgbProgreso(0).Value = 4                        ' Actualizo la barra de progreso
    ' Paso 3 : Verifico el resultado
    If nValidacion% = 0 Then
      MsgBox Choose(gsIdioma, "Validación de Información se completo Correctamente", "Validation of information has completed correctly") & Chr$(13) & Choose(gsIdioma, "Presione Aceptar para Iniciar la Importación de la Información", "You Press Accept to start the import of information"), vbInformation
    ElseIf nValidacion% = 1 Then
      MsgBox Choose(gsIdioma, "Validación de Información se completo con Alertas", "Validation of information has completed with alerts") & Chr$(13) & Choose(gsIdioma, "Presione Aceptar para Imprimir Reporte de Validación y visualizar las Alertas", "You Press Accept to print report of validation and view alerts"), vbExclamation
    Else
      MsgBox Choose(gsIdioma, "Validación de Información tiene Errores", "Validation of information has errors") & Chr$(13) & Choose(gsIdioma, "Presione Aceptar para Imprimir Reporte de Validación y pueda corregir sus Errores", "You Press Accept to print report of validation and can correct errors"), vbCritical
    End If
        
    ' Realizo la transferencia de información
    If nValidacion% = 0 Then
      If MsgBox(Choose(gsIdioma, "Realizamos la Importación de la Información ?" & Chr(13) & IIf(chkPeriodos.Value = vbChecked, "Hasta el Periodo : ", "Del Periodo : "), "It Makes import of information ?" & Chr(13) & IIf(chkPeriodos.Value = vbChecked, "Until the Period : ", "From Period : ")) & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1), vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        ppTransfir_Tablas
        pgbProgreso(0).Value = 5                    ' Actualizo la barra de progreso
        ppTransfir_Proceso
      End If
    ElseIf nValidacion% = 1 Then
      If MsgBox(Choose(gsIdioma, "La Validación encontro Alertas que se pueden Obviar; Realizamos la Importación de la Información ?" & Chr(13) & IIf(chkPeriodos.Value = vbChecked, "Hasta el Periodo : ", "Del Periodo : "), "The validation found alerts that can overlook; Does it make import of information ?" & Chr(13) & IIf(chkPeriodos.Value = vbChecked, "Until the Period : ", "From Period : ")) & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1), vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        ppTransfir_Tablas
        pgbProgreso(0).Value = 5                    ' Actualizo la barra de progreso
        ppTransfir_Proceso
      End If
    End If
    pgbProgreso(0).Value = 6                        ' Actualizo la barra de progreso
        
    pocnnMain.CommitTrans                           ' CONFIRMA TRANSACCION.
    If nValidacion% <> 0 Then
      ' Obtengo los registros del reporte
      With porstMRp
        If .State = adStateOpen Then .Close
        .ActiveConnection = pocnnMain
        '.CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT * "
        .Source = .Source & "FROM " & ps_Prefijo & "trptRPTraInf "
        .Open
      End With
      ' Listado de Errores
      gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "Errores o Alertas de la Validación de Información", "Erros or Alerts of the Validation of Information"), Date, True, False, porstMRp
      With frmMain.rptMain
        .ReportFileName = gsRutRpt & "rptLInfVal.rpt"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
      End With
      porstMRp.Close
    End If
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRPTraInf", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptRPTraInf') DROP TABLE #trptRPTraInf")
   Case 1
    ' Validar ejercicio habilitado
    If Trim(txtDato(0).Text) <> "" Then
      If Not ValidaEjercicio(txtDato(0).Text, Right(cboEjercicio.Text, 4)) Then
        MsgBox Choose(gsIdioma, "Ejercicio seleccionado no se encuentra habilitado; Verificar", "Selected Fiscal year is qualified; Verify"), vbInformation
        cmdAceptar.Enabled = True
        cmdSalir.Enabled = True
        cmdSalir.SetFocus
        Exit Sub
      End If
    End If
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
    ' Paso 1: Realizo la trasnferencia de las tablas
    pgbProgreso(0).Max = chkTransTabla.Count
    pgbProgreso(0).Value = pgbProgreso(0).Min
    If Trim(txtDato(0).Text) <> "" Then
      If MsgBox(Choose(gsIdioma, "Estás Seguro de Inicializar Información Tablas?", "Are you sure Initialize Information Masters Tables?"), vbQuestion + vbYesNo) = vbYes Then
        ppInicializa_Tablas
      End If
    Else
      ppExportar_Tablas
    End If
    pgbProgreso(0).Value = chkTransTabla.Count
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
   Case 2
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
    ' Paso 1: Realizo la centralización de tablas
    pgbProgreso(0).Max = chkCentraTabla.Count
    pgbProgreso(0).Value = pgbProgreso(0).Min
    ppCentraliza_Tablas
    pgbProgreso(0).Value = chkCentraTabla.Count
    ' Paso 2: Realizo la centralización de registros
    pgbProgreso(0).Max = chkCentraProceso.Count
    pgbProgreso(0).Value = pgbProgreso(0).Min
    ppCentraliza_Proceso
    pgbProgreso(0).Value = chkCentraProceso.Count
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
  End Select
  MsgBox TEXT_8008, vbInformation
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  pocnnMain.Close
  Set pocnnMain = Nothing
  Exit Sub
  
Err:
  If pocnnMain.State = adStateOpen Then
    pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
    pocnnMain.Close
    Set pocnnMain = Nothing
  End If
  Set porstMRp = Nothing
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Function pfValida_Tablas() As Integer
  Static sSentencia As String, sTabla As String
  Static nContador As Integer
  Static nRegistro As Double, nNumRegistros As Double

  pgbProgreso(1).Max = chkImporTabla.Count
  pgbProgreso(1).Value = pgbProgreso(1).Min
  For nContador = 0 To chkImporTabla.Count - 1
    ' Verifico que se haya seleccionado
    If chkImporTabla(nContador).Value Then
      sTabla = Choose(nContador + 1, "CoDro", "CoCta", "TgAux", "TgTDc", "CoCCo", "TgTcb")
      lblProgreso(1).Caption = Choose(gsIdioma, "Validando Archivo: ", "Validating File: ") & Trim(chkImporTabla(nContador).Caption)
      Select Case nContador
       Case 0
        ' Diario duplicado en la tabla
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '0', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Diario Registrado :', a.CodDro)", "('Diario Registrado :'+a.CodDro)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND a.codemp=b.codemp "
        sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
        sSentencia = sSentencia & "AND a.CodDro=b.CodDro"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0 And pfValida_Tablas <> 2, 1, pfValida_Tablas)
        ' Diario duplicado en la importación
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '1', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Diario Duplicado (veces) :', CodDro, ' - ', COUNT(*))", "('Diario Duplicado (veces) :'+CodDro+' - '+RTrim(COUNT(*)))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "GROUP BY CodDro "
        sSentencia = sSentencia & "HAVING COUNT(*)<>1"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Diario vacio en el archivo
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '2', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Codigo Diario Vacio :', CodDro)", "('Codigo Diario Vacio :'+CodDro)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodDro, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
       Case 1
        ' Cuenta duplicado en la tabla
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '0', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Cuenta Registrado :', a.CodCta)", "('Cuenta Registrado :'+a.CodCta)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND a.codemp=b.codemp "
        sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
        sSentencia = sSentencia & "AND a.CodCta=b.CodCta"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0 And pfValida_Tablas <> 2, 1, pfValida_Tablas)
        ' Cuenta duplicada en la importación
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '1', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Cuenta Duplicada (veces) :', CodCta, ' - ', COUNT(*))", "('Cuenta Duplicada (veces) :'+CodCta+' - '+RTrim(COUNT(*)))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "GROUP BY CodCta "
        sSentencia = sSentencia & "HAVING COUNT(*)<>1"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Cuenta vacia en el archivo
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '2', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Codigo Cuenta Vacia :', CodCta)", "('Codigo Cuenta Vacia :'+CodCta)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodCta, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
       Case 2
        ' Auxiliar duplicado en la tabla
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '0', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Auxiliar Registrado :', a.CodAux)", "('Auxiliar Registrado :'+a.CodAux)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND a.codemp=b.codemp "
        sSentencia = sSentencia & "AND a.CodAux=b.CodAux"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0 And pfValida_Tablas <> 2, 1, pfValida_Tablas)
        ' Auxiliar duplicado en la importación
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '1', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Auxiliar Duplicado (veces) :', CodAux, ' - ', COUNT(*))", "('Auxiliar Duplicado (veces) :'+CodAux+' - '+RTrim(COUNT(*)))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "GROUP BY CodAux "
        sSentencia = sSentencia & "HAVING COUNT(*)<>1 "
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Auxiliar vacio en el archivo
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '2', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Codigo Auxiliar Vacio :', CodAux)", "('Codigo Auxiliar Vacio :'+CodAux)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodAux, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
       Case 3
        ' Tipo de Documento registrado en la tabla
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '0', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Tipo Documento Registrado :', a.CodTDc)", "('Tipo Documento Registrado :'+a.CodTDc)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND a.codemp=b.codemp "
        sSentencia = sSentencia & "AND a.CodTDc=b.CodTDc"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0 And pfValida_Tablas <> 2, 1, pfValida_Tablas)
        ' Tipo de Documento duplicado en la importación
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '1', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Tipo Documento Duplicado (veces) :', CodTDc, ' - ', COUNT(*))", "('Tipo Documento Duplicado (veces) :'+CodTDc+' - '+RTrim(COUNT(*)))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "GROUP BY CodTDc "
        sSentencia = sSentencia & "HAVING COUNT(*)<>1"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Tipo de documento vacio en el archivo
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '2', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Codigo Tipo Documento Vacio :', CodTDc)", "('Codigo Tipo Documento Vacio :'+CodTDc)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodTDc, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Abreviatura de TD vacia
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '3', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Abreviatura TD en Blanco :', CodTDc)", "('Abreviatura TD en Blanco :'+CodTDc)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AbvTDc, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
       Case 4
        ' Centro de costos duplicado en la tabla
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '0', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Centro de Costo Registrado :', a.CodCCo)", "('Centro de Costo Registrado :'+a.CodCCo)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND a.codemp=b.codemp "
        sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
        sSentencia = sSentencia & "AND a.CodCCo=b.CodCCo"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0 And pfValida_Tablas <> 2, 1, pfValida_Tablas)
        ' Centro de Costo duplicado en la importación
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '1', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Centro Costo Duplicado (veces) :', CodCCo, ' - ', COUNT(*))", "('Centro Costo Duplicado (veces) :'+CodCCo+' - '+RTrim(COUNT(*)))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "GROUP BY CodCCo "
        sSentencia = sSentencia & "HAVING COUNT(*)<>1"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Centro de Costo vacio en el archivo
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '2', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Codigo Centro Costo Vacio :', CodCCo)", "('Codigo Centro Costo Vacio :'+CodCCo)") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodCCo, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
       Case 5
        ' Tipo de cambio registrado en la tabla
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '0', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Tipo de Cambio Registrado :', a.FehTCb)", "('Tipo de Cambio Registrado :'+RTrim(a.FehTCb))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND a.codemp=b.codemp "
        sSentencia = sSentencia & "AND a.FehTCb=b.FehTCb"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0 And pfValida_Tablas <> 2, 1, pfValida_Tablas)
        ' Tipo de cambio duplicado en la importación
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & " SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '1', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Tipo de Cambio Duplicado (veces) :', FehTCb, ' - ', COUNT(*))", "('Tipo de Cambio Duplicado (veces) :'+RTrim(FehTCb)+' - '+RTrim(COUNT(*)))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "GROUP BY FehTCb "
        sSentencia = sSentencia & "HAVING COUNT(*)<>1"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
        ' Fecha de TC vacio en el archivo
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & " SELECT DISTINCT " & Trim$(nContador) & ", '" & Trim(chkImporTabla(nContador).Caption) & "', '2', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Fecha de TC Vacio :', FehTCb)", "('Fecha de TC Vacio :'+RTrim(FehTCb))") & ", '123' "
        sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(FehTCb, '')=''"
        pocnnMain.Execute sSentencia, nNumRegistros
        pfValida_Tablas = IIf(nNumRegistros > 0, 2, pfValida_Tablas)
      End Select
    End If
    pgbProgreso(1).Value = nContador + 1
    DoEvents
  Next nContador

End Function
Function pfNumComprobante(ByVal s_Mes As String, ByVal s_Diario As String) As String
' s_Mes             Perio donde  se genera
' s_Diario          Copdigo de diario para generar numero
    
    Dim porstRetorno As ADODB.Recordset
    Dim s_Sentencia As String
    
    s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), '000000') AS cNumMaxCpb "
    s_Sentencia = s_Sentencia & "FROM CoCpbCab "
    s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND MesPvs='" & s_Mes & "' "
    s_Sentencia = s_Sentencia & "AND CodDro='" & s_Diario & "'"
   
    Set porstRetorno = New ADODB.Recordset
    With porstRetorno
      .ActiveConnection = CONNSTRG & gsNomBDS
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
      .Source = s_Sentencia
      .Open
    End With
    pfNumComprobante = Format(porstRetorno!cNumMaxCpb, "000000")
    porstRetorno.Close
    Set porstRetorno = Nothing

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

Function pfTipoCambio(ByVal s_FechaTcb As String, ByVal s_TipoTcb As String) As Double
' s_FechaTcb    Fecha de tipo de cambio
' s_TipoTcb     Tipo de cambio venta o compra
    
  Dim porstRetorno As ADODB.Recordset
  Dim s_Sentencia As String
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(imptcb_" & IIf(s_TipoTcb = TPOTCB_VTA, "vta", "cpr") & ", 1) AS nTipoCambio "
  s_Sentencia = s_Sentencia & "FROM tgtcb "
  s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
  If ps_Plataforma = pSrvMySql Then
  s_Sentencia = s_Sentencia & "AND DATE_FORMAT(fehtcb,'%d/%m/%Y')=DATE_FORMAT('" & Format(s_FechaTcb, "yyyy-mm-dd") & "', '%d/%m/%Y')"
  ElseIf ps_Plataforma = pSrvSql Then
  s_Sentencia = s_Sentencia & "AND CONVERT(smalldatetime, fehtcb, 103)=CONVERT(smalldatetime, '" & Format(s_FechaTcb, "dd/mm/yyyy") & "', 103)"
  End If
  Set porstRetorno = New ADODB.Recordset
  
  With porstRetorno
    .ActiveConnection = CONNSTRG & gsNomBDS
    ' .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With
  pfTipoCambio = CDec(porstRetorno!nTipoCambio)
  porstRetorno.Close
  Set porstRetorno = Nothing

End Function

Private Function pfValida_Proceso(ByVal nValidacion As Integer) As Integer
    Static sSentencia As String, sTabla As String
    Static nContador As Integer
    Static nRegistro As Double, nNumRegistros As Double

    pfValida_Proceso = nValidacion
    pgbProgreso(1).Max = chkImporProceso.Count
    pgbProgreso(1).Value = pgbProgreso(1).Min
    For nContador = 0 To chkImporProceso.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporProceso(nContador).Value Then
        sTabla = Choose(nContador + 1, "CoCprDoc", "CoVtaDoc", "CoHprDoc", "CoCpbDet")
        lblProgreso(1).Caption = Choose(gsIdioma, "Validando Archivo: ", "Validating File: ") & Trim(chkImporProceso(nContador).Caption)
        
        Select Case nContador
          Case 0
            ' Documento de compras duplicado en la tabla
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '0', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Compra Registrado :', a.CodAux, '/', a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Documento Compra Registrado :'+a.CodAux+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND a.codemp=b.codemp "
            sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
            sSentencia = sSentencia & "AND a.CodAux=b.CodAux "
            sSentencia = sSentencia & "AND a.CodTDc=b.CodTDc "
            sSentencia = sSentencia & "AND a.SerDoc=b.SerDoc "
            sSentencia = sSentencia & "AND a.NroDoc=b.NroDoc "
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = sSentencia & "AND a.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND a.MesPvs=b.MesPvs"
            End If
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Documento de compras duplicado en la importacion
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '1', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Compra Duplicado (veces) :', CodAux, '/', CodTDc, '/', SerDoc,'-',NroDoc, '/', COUNT(*))", "('Documento Compra Duplicado (veces) :'+CodAux+'/'+CodTDc+'/'+SerDoc+'-'+NroDoc+'/'+RTrim(COUNT(*)))") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "GROUP BY CodAux, CodTDc, SerDoc, NroDoc "
            sSentencia = sSentencia & "HAVING COUNT(*)<>1"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Clave de Documento vacio en el archivo
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '2', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Compra vacio :', CodAux, '/', CodTDc, '/', SerDoc,'-',NroDoc)", "('Documento Compra vacio :'+CodAux+'/'+CodTDc+'/'+SerDoc+'-'+NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodAux, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodTDc, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SerDoc, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(NroDoc, '')='')"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Auxiliar no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '3', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Auxiliar (Proveedor) no Registrado :', a.CodAux, '/',a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Auxiliar (Proveedor) no Registrado :'+a.CodAux+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM TgAux b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
            sSentencia = sSentencia & "AND b.IndPrv=" & INDAUX_PRV_ACT & ")"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Diario no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '4', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Diario no Registrado :', a.CodDro, '/',a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Diario no Registrado :'+a.CodDro+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodDro, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoDro b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Valido los procesos adicionales
            If aValidar(nContador, 1) = NvlUsr_Sup Then
              
              
              If aValidar(nContador, 2) = NvlUsr_Sup Then
              End If
            End If
          Case 1
            ' Documento de ventas duplicado en la tabla
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '0', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Ventas Registrado :', a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Documento Ventas Registrado :'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND a.codemp=b.codemp "
            sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
            sSentencia = sSentencia & "AND a.CodTDc=b.CodTDc "
            sSentencia = sSentencia & "AND a.SerDoc=b.SerDoc "
            sSentencia = sSentencia & "AND a.NroDoc=b.NroDoc "
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = sSentencia & "AND b.MesPvs=a.MesPvs "
            End If
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Documento de ventas duplicado en la importacion
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '1', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Ventas Duplicado (veces) :', CodTDc, '/', SerDoc,'-',NroDoc, '/', COUNT(*))", "('Documento Ventas Duplicado (veces) :'+CodTDc+'/'+SerDoc+'-'+NroDoc+'/'+COUNT(*))") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "GROUP BY CodTDc, SerDoc, NroDoc "
            sSentencia = sSentencia & "HAVING COUNT(*)<>1"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Clave de Documento vacio en el archivo
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '2', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Ventas vacio :', CodTDc, '/', SerDoc, '-', NroDoc)", "('Documento Ventas vacio :'+CodTDc+'/'+SerDoc+'-'+NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodTDc, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SerDoc, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(NroDoc, '')='')"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Auxiliar no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '3', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Auxiliar (Cliente) no Registrado :', a.CodAux, '/', a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Auxiliar (Cliente) no Registrado :'+a.CodAux+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM TgAux b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
            sSentencia = sSentencia & "AND b.indcli=" & INDAUX_CLI_ACT & ")"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Diario no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '4', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Diario no Registrado :', a.CodDro, '/', a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Diario no Registrado :'+a.CodDro+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodDro, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoDro b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Tipo de cambio registrado
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '5', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Tipo de Cambio no Registrado :', a.fehope, '/', a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Tipo de Cambio no Registrado :'+a.fehope+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.imptcb, 0.0000)=0.0000"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
'ini 2015-10-13 aumento col detra, const.nro. y fech
            ' detraccion no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '6', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Detraccion no Registrado :', a.tsadetrac, '/', a.CodTDc, '/', a.SerDoc,'-', a.NroDoc)", "('Detraccion :'+a.tsadetrac+'/'+a.CodTDc+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.tsadetrac, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM codetrac b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.coddetrac=a.tsadetrac " & ")"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)

'fin 2015-10-13 aumento col detra, const.nro. y fech
            
          Case 2
            ' Documento de honorarios registrado en la tabla
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '0', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Honorarios Registrado :', a.CodAux, '/', a.SerDoc,'-', a.NroDoc)", "('Documento Honorarios Registrado :'+a.CodAux+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, " & sTabla & " b "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND a.codemp=b.codemp "
            sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
            sSentencia = sSentencia & "AND a.CodAux=b.CodAux "
            sSentencia = sSentencia & "AND a.SerDoc=b.SerDoc "
            sSentencia = sSentencia & "AND a.NroDoc=b.NroDoc "
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = sSentencia & "AND a.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND a.MesPvs=b.MesPvs"
            End If
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
          ' Documento de compras duplicado en la importacion
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '1', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Honorarios Duplicado (veces) :', CodAux, '/', SerDoc,'-',NroDoc, '/', COUNT(*))", "('Documento Honorarios Duplicado (veces) :'+CodAux+'/'+SerDoc+'-'+NroDoc+'/'+RTrim(COUNT(*))") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "GROUP BY CodAux, SerDoc, NroDoc "
            sSentencia = sSentencia & "HAVING COUNT(*)<>1"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Clave de Documento vacio en el archivo
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '2', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Documento Honorarios vacio :', CodAux, '/', SerDoc,'-',NroDoc)", "('Documento Honorarios vacio :'+CodAux+'/'+SerDoc+'-'+NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodAux, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SerDoc, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(NroDoc, '')='')"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Auxiliar no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '3', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Auxiliar (Proveedor) no Registrado :', a.CodAux, '/', a.SerDoc,'-', a.NroDoc)", "('Auxiliar (Proveedor) no Registrado :'+a.CodAux+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM TgAux b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
            sSentencia = sSentencia & "AND b.IndPrv=" & INDAUX_PRV_ACT & " "
            sSentencia = sSentencia & "AND b.TpoPer='" & TPOPER_NAT & "')"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            ' Diario no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '4', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Diario no Registrado :', a.CodDro, '/', a.SerDoc,'-', a.NroDoc)", "('Diario no Registrado :'+a.CodDro+'/'+a.SerDoc+'-'+a.NroDoc)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodDro, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoDro b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
          Case 3
            If chkEliminar(nContador).Value = vbUnchecked Then
              ' Comprobante de diario registrado en la tabla
              sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
              sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '0', "
              sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Comprobante Diario Registrado :', a.MesPvs, '/', a.CodDro,'-', a.NroCpb)", "('Comprobante Diario Registrado :'+a.MesPvs+'/'+a.CodDro+'-'+a.NroCpb)") & ", '123' "
              sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a, CoCpbCab b "
              sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND a.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
              sSentencia = sSentencia & "AND a.codemp=b.codemp "
              sSentencia = sSentencia & "AND a.pdoano=b.pdoano "
              sSentencia = sSentencia & "AND a.MesPvs=b.MesPvs "
              sSentencia = sSentencia & "AND a.CodDro=b.CodDro "
              sSentencia = sSentencia & "AND a.NroCpb=b.NroCpb "
              pocnnMain.Execute sSentencia, nNumRegistros
              pfValida_Proceso = IIf(nNumRegistros > 0 And pfValida_Proceso <> 2, 1, pfValida_Proceso)
            End If
            ' Clave de comprobante de diario vacio en el archivo
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '1', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Comprobante de Diario vacio :', MesPvs, '/', CodDro, '-', NroCpb)", "('Comprobante de Diario vacio :'+MesPvs+'/'+CodDro+'-'+NroCpb)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MesPvs, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodDro, '')='' "
            sSentencia = sSentencia & "OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(NroCpb, '')='')"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Diario no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '2', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Diario de Comprobante no Registrado :', a.MesPvs, '/', a.CodDro,'-', a.NroCpb)", "('Diario de Comprobante no Registrado :'+a.MesPvs+'/'+a.CodDro+'-'+a.NroCpb)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodDro, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoDro b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Cuenta no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '3', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Cuenta Contable no Registrado :', a.Codcta, '/', a.CodDro,'-', a.NroCpb)", "('Cuenta Contable no Registrado :'+a.CodCta+'/'+a.CodDro+'-'+a.NroCpb)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCta, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoCta b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodCta=a.CodCta)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Auxiliar no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '4', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Auxiliar no Registrado :', a.CodAux, '/', a.CodDro,'-', a.NroCpb)", "('Auxiliar no Registrado :'+a.CodAux+'/'+a.CodDro+'-'+a.NroCpb)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM TgAux b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Centro de Costos no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '5', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Centro de Costos no Registrado :', a.CodCCo, '/', a.CodDro,'-', a.NroCpb)", "('Centro de Costos no Registrado :'+a.CodCCo+'/'+a.CodDro+'-'+a.NroCpb)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCCo, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoCCo b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodCCo=a.CodCCo)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
            ' Tipo de Documento no registrado en la tabla general
            sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
            sSentencia = sSentencia & "SELECT DISTINCT " & Trim$(nContador + 5) & ", '" & Trim(chkImporProceso(nContador).Caption) & "', '6', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Tipo de Documento no Registrado :', a.CodTdc, '/', a.CodDro,'-', a.NroCpb)", "('Tipo de Documento no Registrado :'+a.CodTdc+'/'+a.CodDro+'-'+a.NroCpb)") & ", '123' "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTdc, '')<>'' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM TgTDc b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodTdc=a.CodTdc)"
            pocnnMain.Execute sSentencia, nNumRegistros
            pfValida_Proceso = IIf(nNumRegistros > 0, 2, pfValida_Proceso)
        End Select
      End If
      pgbProgreso(1).Value = nContador + 1
      DoEvents
    Next nContador

End Function

Private Sub ppImporta_Proceso()
  Static sSentencia As String
  Static sArchivo As String, sMilinea As String
  Static nContador As Integer, nArchivo As Integer
  Static nColumnas As Integer
  Static nRegistro As Double, nNumRegistros As Double
  Static sRegOrden As String, nRegOrden As Double
  Static aRegistros()
  Static porstTmp As ADODB.Recordset
  Dim aArchivo(), aColumna(), aTabla(), aActualiza()
  Dim sNewComprobante As String, sComprobante As String
  Dim sDiario As String, sTabla As String, sExpresion As String
  Dim nSecuencia As Integer
  Dim nTipoCambio As Double, nImporte As Double
  Dim sPeriodoIn As String

  ReDim aValidar(2, 2)
  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  ' Importo las tablas de acuerdo a la selección
  nArchivo = FreeFile
  For nContador = 0 To chkImporProceso.Count - 1
    ' Verifico que se haya seleccionado
    If chkImporProceso(nContador).Value Then
      ' Abro Archivo de Texto
      aArchivo = Choose(nContador + 1, Array("rc", "cc", "co"), Array("rv", "vc", "vo"), Array("rh", "hc", "ho"), Array("rd"))
      aTabla = Choose(nContador + 1, Array("CoCprDoc", "cocprdoccta", "cocprdoccco"), Array("covtadoc", "covtadoccta", "covtadoccco"), Array("CoHprDoc", "cohprdoccta", "cohprdoccco"), Array("CoCpbDet"))
      
      'MsgBox aArchivo(nContador)
      'MsgBox aTabla(nContador)
      
      '2015-10-09 aumento col detra, const.nro. y fech aColumna = Choose(nContador + 1, Array(44, 13, 11), Array(39, 14, 10), Array(30, 12, 10), Array(30))
      '2016-03-16 aumento vta=codmon,tpo,ser,doc ref  aColumna = Choose(nContador + 1, Array(44, 13, 11), Array(40, 14, 10), Array(30, 12, 10), Array(30))
      '2016-05-17 adiciona ple mon y bns aColumna = Choose(nContador + 1, Array(44, 13, 11), Array(44, 14, 10), Array(30, 12, 10), Array(30))
      aColumna = Choose(nContador + 1, Array(46, 13, 11), Array(44, 14, 10), Array(30, 12, 10), Array(30))
      aActualiza = Choose(nContador + 1, Array(INDPREGEN_INA, INDPREGEN_INA, INDPREGEN_INA), Array(INDPREGEN_INA, INDPREGEN_INA, INDPREGEN_INA), Array(INDPREGEN_INA, INDPREGEN_INA, INDPREGEN_INA), Array(INDPREGEN_INA))
      sPeriodoIn = Choose(nContador + 1, gsMesAct, gsMesAct, gsMesAct, IIf(chkPeriodos.Value = vbChecked, "", gsMesAct))
      
      nSecuencia = 0
      sArchivo = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & aArchivo(nSecuencia) & gsAnoAct & sPeriodoIn & ".txt"
      ' Desactivo la opcion si no existe archivo
      chkImporProceso(nContador).Value = vbUnchecked
      If Dir$(sArchivo, vbNormal) <> "" Then
        ' Activo la opcion si existe archivo
        chkImporProceso(nContador).Value = vbChecked
        ' Recorro si teiene elementos
        For nSecuencia = 0 To UBound(aArchivo)
          sArchivo = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & aArchivo(nSecuencia) & gsAnoAct & sPeriodoIn & ".txt"
          nColumnas = aColumna(nSecuencia)
          If Dir$(sArchivo, vbNormal) <> "" Then
            Open sArchivo For Input As #nArchivo
            nNumRegistros = gfRedond(LOF(nArchivo), 0)
            If nNumRegistros > 0 Then
              pgbProgreso(1).Max = nNumRegistros
              pgbProgreso(1).Value = pgbProgreso(1).Min
              ' Activo la validacion de procesos
              If nContador <= 2 Then
                aValidar(nContador, nSecuencia) = NvlUsr_Sup
              End If
              ' Elimino y genero el archivo y abro el recordset temporal
              sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & Len(aTabla(nSecuencia)) + 5 & ")='#tmp" & aTabla(nSecuencia) & "_') DROP TABLE #tmp" & aTabla(nSecuencia)
              pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp" & aTabla(nSecuencia), sSentencia)
              
              sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmp" & aTabla(nSecuencia) & " ", "")
              sSentencia = sSentencia & "SELECT * "
              If nContador = 3 Then
                sSentencia = sSentencia & ", gloite AS glocpb, gloitex AS glocpbx, 0 AS IndNCu, 0 AS IndAnu "
              End If
              sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & ps_Prefijo & "tmp" & aTabla(nSecuencia) & " ", "")
              sSentencia = sSentencia & "FROM " & aTabla(nSecuencia) & " "
              sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND UsrCre='sysmavm'"
              pocnnMain.Execute sSentencia
              With porstTmp
                If .State = adStateOpen Then .Close
                .ActiveConnection = pocnnMain
                .Source = "SELECT * "
                .Source = .Source & "FROM " & ps_Prefijo & "tmp" & aTabla(nSecuencia)
                .CursorType = adOpenDynamic
                .LockType = adLockOptimistic
                .Open
              End With
              ' Inicializo el numero de comprobante
              sRegOrden = "": sDiario = ""
              sComprobante = "": sNewComprobante = ""
              ' Verifico activacion de archivo
              aActualiza(nSecuencia) = INDPREGEN_ACT
              
              ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
              ReDim aRegistros(nColumnas)
              lblProgreso(1).Caption = Choose(gsIdioma, "Importando Archivo: ", "importing File: ") & Trim(chkImporProceso(nContador).Caption) & " (" & gsRUCEmp & aArchivo(nSecuencia) & gsAnoAct & sPeriodoIn & ")"
              Do While Not EOF(nArchivo)
                Line Input #nArchivo, sMilinea
                nRegistro = nRegistro + 1
                pRegistro_Texto sMilinea, nColumnas, aRegistros
                With porstTmp
                  .AddNew
                  !codemp = gsCodEmp
                  Select Case nContador
                   Case 0 'registro compras
                    If nSecuencia = 0 Then
                      ' Obtengo numero de comprobante
                      If sRegOrden <> Left(Trim$(aRegistros(18)), 4) & Left(Trim$(aRegistros(19)), 6) Then
                        sRegOrden = Left(aRegistros(18), 4) & Left(aRegistros(19), 6)
                        If sDiario <> Left(Trim$(aRegistros(18)), 4) Then
                          sDiario = Left(Trim$(aRegistros(18)), 4)
                          sNewComprobante = pfNumComprobante(Left(aRegistros(17), 2), sDiario)
                        End If
                        sComprobante = Left(aRegistros(19), 6)
                        If chkCorrelativo.Value = vbChecked Then
                          sNewComprobante = Format(Val(sNewComprobante) + 1, "000000")
                          sComprobante = sNewComprobante
                        End If
                      End If
                      ' Obtengo el tipo de cambio de la fecha
                      
                      ' Informacion del registro
                      !pdoano = aRegistros(1)
                      !codaux = aRegistros(2)
                      !codtdc = aRegistros(3)
                      !serdoc = aRegistros(4)
                      !nrodoc = aRegistros(5)
                      !fehope = IIf(IsDate(aRegistros(6)), Format(aRegistros(6), "dd/mm/yyyy"), Null)
                      !feedoc = IIf(IsDate(aRegistros(7)), Format(aRegistros(7), "dd/mm/yyyy"), Null)
                      !fevdoc = IIf(IsDate(aRegistros(8)), Format(aRegistros(8), "dd/mm/yyyy"), Null)
                      !ferdoc = IIf(IsDate(aRegistros(9)), Format(aRegistros(9), "dd/mm/yyyy"), Null)
                      !tpomon = IIf(aRegistros(10) = "N", TPOMON_NAC, TPOMON_EXT)
                      !ImpTCb = CDec(Format(Val(aRegistros(11)), FORMATO_NUM_2))
                      !PctIGV = CDec(Format(Val(aRegistros(12)), FORMATO_NUM_1))
                      !PctISC = CDec(Format(Val(aRegistros(13)), FORMATO_NUM_1))
                      !RefDoc = aRegistros(14)
                      sExpresion = Left(aRegistros(15), 50)
                      !GloDoc = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(16), 50)
                      !glodocx = IIf(sExpresion = "", Null, sExpresion)
                      !mespvs = aRegistros(17)
                      !coddro = IIf(sDiario = "", Null, sDiario)
                      !NroCpb = sComprobante
                      !indcdt = IIf(Trim$(aRegistros(20)) <> "", INDCDT_ACT, INDCDT_INA)
                      !NroCDt = aRegistros(20)
                      !FehCDt = IIf(aRegistros(21) = "", Null, Format(aRegistros(21), "dd/mm/yyyy"))
                      
                      'MONEDA NACIONAL
                      !impogr_mn = CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1))
                      !ImpOGN_MN = CDec(Format(Val(aRegistros(23)), FORMATO_NUM_1))
                      !ImpONG_MN = CDec(Format(Val(aRegistros(24)), FORMATO_NUM_1))
                      !impexo_mn = CDec(Format(Val(aRegistros(25)), FORMATO_NUM_1))
                      !impigv_mn = CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1))
                      '!ImpIGV_OGr_MN = CDec(Format(Val(aRegistros(27)), FORMATO_NUM_1))
                      '!ImpIGV_OGN_MN = CDec(Format(Val(aRegistros(28)), FORMATO_NUM_1))
                      '!ImpIGV_ONG_MN = CDec(Format(Val(aRegistros(29)), FORMATO_NUM_1))
                      !ImpIGV_OGr_MN = IIf(CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1)) > 0, CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1)), 0)
                      !ImpIGV_OGN_MN = IIf(CDec(Format(Val(aRegistros(23)), FORMATO_NUM_1)) > 0, CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1)), 0)
                      !ImpIGV_ONG_MN = IIf(CDec(Format(Val(aRegistros(24)), FORMATO_NUM_1)) > 0, CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1)), 0)
                      !impisc_mn = CDec(Format(Val(aRegistros(30)), FORMATO_NUM_1))
                      !impoim_mn = CDec(Format(Val(aRegistros(31)), FORMATO_NUM_1))
                      !imptot_mn = CDec(Format(Val(aRegistros(32)), FORMATO_NUM_1))
                      'MONEDA EXTRANJERA
                      !impogr_me = CDec(Format(Val(aRegistros(33)), FORMATO_NUM_1))
                      !ImpOGN_ME = CDec(Format(Val(aRegistros(34)), FORMATO_NUM_1))
                      !ImpONG_ME = CDec(Format(Val(aRegistros(35)), FORMATO_NUM_1))
                      !impexo_me = CDec(Format(Val(aRegistros(36)), FORMATO_NUM_1))
                      !impigv_me = CDec(Format(Val(aRegistros(37)), FORMATO_NUM_1))
                      '!ImpIGV_OGr_ME = CDec(Format(Val(aRegistros(38)), FORMATO_NUM_1))
                      '!ImpIGV_OGN_ME = CDec(Format(Val(aRegistros(39)), FORMATO_NUM_1))
                      '!ImpIGV_ONG_ME = CDec(Format(Val(aRegistros(40)), FORMATO_NUM_1))
                      !ImpIGV_OGr_ME = IIf(CDec(Format(Val(aRegistros(33)), FORMATO_NUM_1)) > 0, CDec(Format(Val(aRegistros(37)), FORMATO_NUM_1)), 0)
                      !ImpIGV_OGN_ME = IIf(CDec(Format(Val(aRegistros(34)), FORMATO_NUM_1)) > 0, CDec(Format(Val(aRegistros(37)), FORMATO_NUM_1)), 0)
                      !ImpIGV_ONG_ME = IIf(CDec(Format(Val(aRegistros(35)), FORMATO_NUM_1)) > 0, CDec(Format(Val(aRegistros(37)), FORMATO_NUM_1)), 0)
                      !impisc_me = CDec(Format(Val(aRegistros(41)), FORMATO_NUM_1))
                      !impoim_me = CDec(Format(Val(aRegistros(42)), FORMATO_NUM_1))
                      !imptot_me = CDec(Format(Val(aRegistros(43)), FORMATO_NUM_1))
                      
                      'modificado 03/09/2013 RC
                      !feedoc_ref = IIf(IsDate(aRegistros(7)), Format(aRegistros(7), "dd/mm/yyyy"), Null)

                      
                      !IndAnu = IIf(aRegistros(44) = "S", INDPREGEN_ACT, INDPREGEN_INA)
                      !indpregen = INDPREGEN_INA
                      ' Se agreo default
                      !indgen = INDPREGEN_INA
                      '!indcta_ogr = IIf(CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1)) > 0, 2, 0)
                      '!IndCta_OGN = IIf(CDec(Format(Val(aRegistros(23)), FORMATO_NUM_1)) > 0, 2, 0)
                      '!IndCta_ONG = IIf(CDec(Format(Val(aRegistros(24)), FORMATO_NUM_1)) > 0, 2, 0)
                      '!indcta_exo = IIf(CDec(Format(Val(aRegistros(25)), FORMATO_NUM_1)) > 0, 2, 0)
                      '!indcta_isc = INDPREGEN_INA
                      '!indcta_igv = IIf(CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1)) > 0, 2, 0)
                      '!indcta_oim = IIf(CDec(Format(Val(aRegistros(31)), FORMATO_NUM_1)) > 0, 2, 0)
                      '!indcta_tot = IIf(CDec(Format(Val(aRegistros(32)), FORMATO_NUM_1)) > 0, 2, 0)
                      
                       ' cambio de flag TC
                       !indcta_ogr = INDPREGEN_ACT
                       !IndCta_OGN = INDPREGEN_ACT
                       !IndCta_ONG = INDPREGEN_ACT
                       !indcta_exo = INDPREGEN_ACT
                       !indcta_isc = INDPREGEN_ACT
                       !indcta_igv = INDPREGEN_ACT
                       !indcta_oim = INDPREGEN_ACT
                       !indcta_tot = INDPREGEN_ACT
                      
'ini 2016-05-17 adiciona ple mon y bns
                      !tpobns = IIf(aRegistros(45) = "", Null, aRegistros(45))
                      !codmon = IIf(aRegistros(46) = "", Null, aRegistros(46))
'fin 2016-05-17 adiciona ple mon y bns

                    ElseIf nSecuencia = 1 Then
                      !pdoano = aRegistros(1)
                      !codaux = aRegistros(2)
                      !codtdc = aRegistros(3)
                      !serdoc = aRegistros(4)
                      !nrodoc = aRegistros(5)
                      !tpocnc = aRegistros(6)
                      !orden = aRegistros(7)
                      !codcta = aRegistros(8)
                      sExpresion = Left(aRegistros(9), 50)
                      !glodet = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(10), 50)
                      !glodetx = IIf(sExpresion = "", Null, sExpresion)
                      !codruc = aRegistros(11)
                      !impcta_mn = CDec(Format(Val(aRegistros(12)), FORMATO_NUM_1))
                      !impcta_me = CDec(Format(Val(aRegistros(13)), FORMATO_NUM_1))
                    ElseIf nSecuencia = 2 Then
                      !pdoano = aRegistros(1)
                      !codaux = aRegistros(2)
                      !codtdc = aRegistros(3)
                      !serdoc = aRegistros(4)
                      !nrodoc = aRegistros(5)
                      !tpocnc = aRegistros(6)
                      !orden = aRegistros(7)
                      !codcta = aRegistros(8)
                      !codcco = aRegistros(9)
                      !impcco_mn = CDec(Format(Val(aRegistros(10)), FORMATO_NUM_1))
                      !impcco_me = CDec(Format(Val(aRegistros(11)), FORMATO_NUM_1))
                    End If
                   Case 1 'registro ventas
                    If nSecuencia = 0 Then
                      ' Obtengo numero de comprobante
                      If sRegOrden <> Left(Trim$(aRegistros(19)), 4) & Left(Trim$(aRegistros(20)), 6) Then
                        sRegOrden = Left(aRegistros(19), 4) & Left(aRegistros(20), 6)
                        If sDiario <> Left(Trim$(aRegistros(19)), 4) Then
                          sDiario = Left(Trim$(aRegistros(19)), 4)
                          sNewComprobante = pfNumComprobante(Left(aRegistros(18), 2), sDiario)
                        End If
                        sComprobante = Left(aRegistros(20), 6)
                        If chkCorrelativo.Value = vbChecked Then
                          sNewComprobante = Format(Val(sNewComprobante) + 1, "000000")
                          sComprobante = sNewComprobante
                        End If
                      End If
                      ' Información del registro
                      !pdoano = aRegistros(1)
                      !codtdc = aRegistros(2)
                      !serdoc = aRegistros(3)
                      !nrodoc = aRegistros(4)
                      !fehope = IIf(IsDate(aRegistros(5)), Format(aRegistros(5), "dd/mm/yyyy"), Null)
                      !SerDoc_Fin = aRegistros(6)
                      !nrodoc_fin = aRegistros(7)
                      !codaux = IIf(Trim$(aRegistros(8)) = "", Null, aRegistros(8))
                      !feedoc = IIf(IsDate(aRegistros(9)), Format(aRegistros(9), "dd/mm/yyyy"), Null)
                      !fevdoc = IIf(IsDate(aRegistros(10)), Format(aRegistros(10), "dd/mm/yyyy"), Null)
                      !tpomon = IIf(aRegistros(11) = "N", TPOMON_NAC, TPOMON_EXT)
                      ' Obtengo el tipo de cambio de la fecha si es cero
                      nTipoCambio = CDec(Format(Val(aRegistros(12)), FORMATO_NUM_2))
                      !ImpTCb = nTipoCambio
                      !PctIGV = CDec(Format(Val(aRegistros(13)), FORMATO_NUM_1))
                      !PctISC = CDec(Format(Val(aRegistros(14)), FORMATO_NUM_1))
                      !RefDoc = aRegistros(15)
                      !TpoGlo_Rtc = INDPREGEN_INA
                      sExpresion = Left(aRegistros(16), 50)
                      !GloDoc = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(17), 50)
                      !glodocx = IIf(sExpresion = "", Null, sExpresion)
                      !mespvs = aRegistros(18)
                      !coddro = IIf(sDiario = "", Null, sDiario)
                      !NroCpb = sComprobante
                      ' Asignacion de importes moneda nacional
                      !impogr_mn = CDec(Format(Val(aRegistros(21)), FORMATO_NUM_1))
                      !ImpExp_mn = CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1))
                      !impexo_mn = CDec(Format(Val(aRegistros(23)), FORMATO_NUM_1))
                      !impigv_mn = CDec(Format(Val(aRegistros(24)), FORMATO_NUM_1))
                      !impisc_mn = CDec(Format(Val(aRegistros(25)), FORMATO_NUM_1))
                      !impoim_mn = CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1))
                      !imptot_mn = CDec(Format(Val(aRegistros(27)), FORMATO_NUM_1))
                      ' Asignacion de importes moneda extranjera
                      !impogr_me = CDec(Format(Val(aRegistros(28)), FORMATO_NUM_1))
                      !impexp_me = CDec(Format(Val(aRegistros(29)), FORMATO_NUM_1))
                      !impexo_me = CDec(Format(Val(aRegistros(30)), FORMATO_NUM_1))
                      !impigv_me = CDec(Format(Val(aRegistros(31)), FORMATO_NUM_1))
                      !impisc_me = CDec(Format(Val(aRegistros(32)), FORMATO_NUM_1))
                      !impoim_me = CDec(Format(Val(aRegistros(33)), FORMATO_NUM_1))
                      !imptot_me = CDec(Format(Val(aRegistros(34)), FORMATO_NUM_1))
                      
                      'modificado 03/09/2013 RV
                      !feedoc_ref = IIf(IsDate(aRegistros(9)), Format(aRegistros(9), "dd/mm/yyyy"), Null)

                      
                      !IndAnu = IIf(aRegistros(35) = "S", INDPREGEN_ACT, INDPREGEN_INA)
                      !indpregen = INDPREGEN_INA
                      ' Se agreo default
                      !indgen = INDPREGEN_INA
                      
'ini 2015-10-09 aumento col detra, const.nro. y fech
                      '!indcdt = aRegistros(36)
                      !IndAnu = IIf(Trim(aRegistros(37)) = "", INDCONS_DPO_0, INDCONS_DPO_1)
                      !FehCDt = IIf(IsDate(aRegistros(36)), Format(aRegistros(36), "dd/mm/yyyy"), Null)
                      !NroCDt = aRegistros(37)
                      !tsadetrac = aRegistros(38)
                      !pctdetrac = aRegistros(39)
'fin 2015-10-09 aumento col detra, const.nro. y fech
'ini 2016-03-16 inlcuye codmon ,tpo,ser,doc ref en cabeza venta
                      !codmon = aRegistros(40)
                      !codtdc_ref = IIf(aRegistros(41) = "", Null, aRegistros(41))
                      !serdoc_ref = IIf(aRegistros(42) = "", Null, aRegistros(41))
                      !nrodoc_ref = IIf(aRegistros(43) = "", Null, aRegistros(41))
'fin 2016-03-16 inlcuye codmon ,tpo,ser,doc ref en cabeza venta
                      
                      
                      If aRegistros(2) = "07" Then
                        !indcta_ogr = IIf(CDec(Format(Val(aRegistros(21)), FORMATO_NUM_1)) > 0, 2, 0)
                        !indcta_exp = IIf(CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1)) > 0, 2, 0)
                        !indcta_exo = IIf(CDec(Format(Val(aRegistros(23)), FORMATO_NUM_1)) > 0, 2, 0)
                        !indcta_igv = IIf(CDec(Format(Val(aRegistros(24)), FORMATO_NUM_1)) > 0, 2, 0)
                        !indcta_isc = IIf(CDec(Format(Val(aRegistros(25)), FORMATO_NUM_1)) > 0, 2, 0)
                        !indcta_oim = IIf(CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1)) > 0, 2, 0)
                        !indcta_tot = IIf(CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1)) > 0, 2, 0)
                      Else
                        !indcta_ogr = INDPREGEN_INA
                        !indcta_exp = INDPREGEN_INA
                        !indcta_exo = INDPREGEN_INA
                        !indcta_igv = INDPREGEN_INA
                        !indcta_isc = INDPREGEN_INA
                        !indcta_oim = INDPREGEN_INA
                        !indcta_tot = INDPREGEN_INA
                      End If
                      ' datos documento afecta
                      '!codtdc_ref = IIf(Trim$(aRegistros(36)) = "", Null, aRegistros(36))
                      '!serdoc_ref = IIf(Trim$(aRegistros(37)) = "", Null, aRegistros(37))
                      '!nrodoc_ref = IIf(Trim$(aRegistros(38)) = "", Null, aRegistros(38))
                      '!feedoc_ref = IIf(IsDate(aRegistros(39)), Format(aRegistros(39), "dd/mm/yyyy"), Null)
                    ElseIf nSecuencia = 1 Then
                      !pdoano = aRegistros(1)
                      !codtdc = aRegistros(2)
                      !serdoc = aRegistros(3)
                      !nrodoc = aRegistros(4)
                      !tpocnc = aRegistros(5)
                      !orden = aRegistros(6)
                      !codcta = aRegistros(7)
                      sExpresion = Left(aRegistros(8), 250)
                      !glodet0 = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(9), 250)
                      !glodet1 = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(10), 250)
                      !glodet0x = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(11), 250)
                      !glodet1x = IIf(sExpresion = "", Null, sExpresion)
                      !codruc = aRegistros(12)
                      !impcta_mn = CDec(Format(Val(aRegistros(13)), FORMATO_NUM_1))
                      !impcta_me = CDec(Format(Val(aRegistros(14)), FORMATO_NUM_1))
                    ElseIf nSecuencia = 2 Then
                      !pdoano = aRegistros(1)
                      !codtdc = aRegistros(2)
                      !serdoc = aRegistros(3)
                      !nrodoc = aRegistros(4)
                      !tpocnc = aRegistros(5)
                      !orden = aRegistros(6)
                      !codcta = aRegistros(7)
                      !codcco = aRegistros(8)
                      !impcco_mn = CDec(Format(Val(aRegistros(9)), FORMATO_NUM_1))
                      !impcco_me = CDec(Format(Val(aRegistros(10)), FORMATO_NUM_1))
                    End If
                   Case 2 'registro honorarios
                    If nSecuencia = 0 Then
                      ' Obtengo numero de comprobante
                      If sRegOrden <> Left(Trim$(aRegistros(15)), 4) & Left(Trim$(aRegistros(16)), 6) Then
                        sRegOrden = Left(aRegistros(15), 4) & Left(aRegistros(16), 6)
                        If sDiario <> Left(Trim$(aRegistros(15)), 4) Then
                          sDiario = Left(Trim$(aRegistros(15)), 4)
                          sNewComprobante = pfNumComprobante(Left(aRegistros(14), 2), sDiario)
                        End If
                        sComprobante = Left(aRegistros(16), 6)
                        If chkCorrelativo.Value = vbChecked Then
                          sNewComprobante = Format(Val(sNewComprobante) + 1, "000000")
                          sComprobante = sNewComprobante
                        End If
                      End If
                      ' Información del registro
                      !pdoano = aRegistros(1)
                      !codaux = aRegistros(2)
                      !serdoc = aRegistros(3)
                      !nrodoc = aRegistros(4)
                      !fehope = IIf(IsDate(aRegistros(5)), Format(aRegistros(5), "dd/mm/yyyy"), Null)
                      !feedoc = IIf(IsDate(aRegistros(6)), Format(aRegistros(6), "dd/mm/yyyy"), Null)
                      !tpomon = IIf(aRegistros(7) = "N", TPOMON_NAC, TPOMON_EXT)
                      !ImpTCb = CDec(Format(Val(aRegistros(8)), FORMATO_NUM_2))
                      !PctIR4 = CDec(Format(Val(aRegistros(9)), FORMATO_NUM_1))
                      !PctIES = CDec(Format(Val(aRegistros(10)), FORMATO_NUM_1))
                      !RefDoc = aRegistros(11)
                      sExpresion = Left(aRegistros(12), 50)
                      !GloDoc = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(13), 50)
                      !glodocx = IIf(sExpresion = "", Null, sExpresion)
                      !mespvs = aRegistros(14)
                      !coddro = IIf(sDiario = "", Null, sDiario)
                      !NroCpb = sComprobante
                      !ImpBru_MN = CDec(Format(Val(aRegistros(17)), FORMATO_NUM_1))
                      !ImpIR4_MN = CDec(Format(Val(aRegistros(18)), FORMATO_NUM_1))
                      !ImpIES_MN = CDec(Format(Val(aRegistros(19)), FORMATO_NUM_1))
                      !ImpORt_MN = CDec(Format(Val(aRegistros(20)), FORMATO_NUM_1))
                      !ImpNet_MN = CDec(Format(Val(aRegistros(21)), FORMATO_NUM_1))
                      !ImpBru_ME = CDec(Format(Val(aRegistros(22)), FORMATO_NUM_1))
                      !ImpIR4_ME = CDec(Format(Val(aRegistros(23)), FORMATO_NUM_1))
                      !ImpIES_ME = CDec(Format(Val(aRegistros(24)), FORMATO_NUM_1))
                      !ImpORt_ME = CDec(Format(Val(aRegistros(25)), FORMATO_NUM_1))
                      !ImpNet_ME = CDec(Format(Val(aRegistros(26)), FORMATO_NUM_1))
                      
                      !IndAfeIR4 = IIf(aRegistros(27) = "S", INDPREGEN_ACT, INDPREGEN_INA)
                      !IndAfeIES = IIf(aRegistros(28) = "S", INDPREGEN_ACT, INDPREGEN_INA)
                      !IndAfeORt = IIf(aRegistros(29) = "S", INDPREGEN_ACT, INDPREGEN_INA)
                      
                      !IndAnu = IIf(aRegistros(30) = "S", INDPREGEN_ACT, INDPREGEN_INA)
                      !indpregen = INDPREGEN_INA
                      
                      '28/10/2008
                      !IndCta_Bru = 2
                      !IndCta_Net = 2
                      
                    ElseIf nSecuencia = 1 Then
                      !pdoano = aRegistros(1)
                      !codaux = aRegistros(2)
                      !serdoc = aRegistros(3)
                      !nrodoc = aRegistros(4)
                      !tpocnc = aRegistros(5)
                      !orden = aRegistros(6)
                      !codcta = aRegistros(7)
                      sExpresion = Left(aRegistros(8), 50)
                      !glodet = IIf(sExpresion = "", Null, sExpresion)
                      sExpresion = Left(aRegistros(9), 50)
                      !glodetx = IIf(sExpresion = "", Null, sExpresion)
                      !codruc = aRegistros(10)
                      !impcta_mn = CDec(Format(Val(aRegistros(11)), FORMATO_NUM_1))
                      !impcta_me = CDec(Format(Val(aRegistros(12)), FORMATO_NUM_1))
                    ElseIf nSecuencia = 2 Then
                      !pdoano = aRegistros(1)
                      !codaux = aRegistros(2)
                      !serdoc = aRegistros(3)
                      !nrodoc = aRegistros(4)
                      !tpocnc = aRegistros(5)
                      !orden = aRegistros(6)
                      !codcta = aRegistros(7)
                      !codcco = aRegistros(8)
                      !impcco_mn = CDec(Format(Val(aRegistros(9)), FORMATO_NUM_1))
                      !impcco_me = CDec(Format(Val(aRegistros(10)), FORMATO_NUM_1))
                    End If
                   Case 3 ' registro diario
                    If sRegOrden <> Left(aRegistros(2), 4) & Left(aRegistros(3), 6) Then
                      sRegOrden = Left(aRegistros(2), 4) & Left(aRegistros(3), 6)
                      If sDiario <> Left(aRegistros(2), 4) Then
                        sDiario = Left(aRegistros(2), 4)
                        sNewComprobante = pfNumComprobante(Left(aRegistros(4), 2), sDiario)
                      End If
                      sComprobante = Left(aRegistros(3), 6)
                      If chkCorrelativo.Value = vbChecked Then
                        sNewComprobante = Format(Val(sNewComprobante) + 1, "000000")
                        sComprobante = sNewComprobante
                      End If
                      nRegOrden = 0
                    End If
                    
                    nRegOrden = nRegOrden + 1
                    !pdoano = aRegistros(1)
                    !mespvs = Left(aRegistros(4), 2)
                    !coddro = sDiario
                    !NroCpb = sComprobante
                    !fehope = Format(aRegistros(5), "dd/mm/yyyy")
                    sExpresion = Left(aRegistros(6), 50)
                    !glocpb = IIf(sExpresion = "", Null, sExpresion)
                    sExpresion = Left(aRegistros(7), 50)
                    !glocpbx = IIf(sExpresion = "", Null, sExpresion)
                    !tpognr = IIf(aRegistros(8) = "C", TPOGNR_CPR, IIf(aRegistros(8) = "V", TPOGNR_VTA, IIf(aRegistros(8) = "H", TPOGNR_HPR, IIf(aRegistros(8) = "D", TPOGNR_DRO, IIf(aRegistros(8) = "A", TPOGNR_DCA, TPOGNR_DST)))))
                    !NroIte = nRegOrden
                    !blqite = aRegistros(10)
                    !codtdc = IIf(aRegistros(11) = "", Null, aRegistros(11))
                    !codcta = Left(aRegistros(12), 8)
                    !codcco = IIf(aRegistros(13) = "", Null, aRegistros(13))
                    !codaux = IIf(aRegistros(14) = "", Null, aRegistros(14))
                    !serdoc = aRegistros(15)
                    !nrodoc = Left(aRegistros(16), 10)
                    !feedoc = IIf(IsDate(aRegistros(17)), Format(aRegistros(17), "dd/mm/yyyy"), Null)
                    !fevdoc = IIf(IsDate(aRegistros(18)), Format(aRegistros(18), "dd/mm/yyyy"), Null)
                    !ferdoc = IIf(IsDate(aRegistros(19)), Format(aRegistros(19), "dd/mm/yyyy"), Null)
                    !RefDoc = Left(aRegistros(20), 20)
                    sExpresion = Left(aRegistros(21), 60)
                    !GloIte = IIf(sExpresion = "", Null, sExpresion)
                    sExpresion = Left(aRegistros(22), 60)
                    !GloItex = IIf(sExpresion = "", Null, sExpresion)
                    !TpoCtb = IIf(aRegistros(23) = "D", TPOCTB_DEB, TPOCTB_HAB)
                    !TpoPvs = IIf(aRegistros(24) = "P", TPOPVS_PVS, IIf(aRegistros(24) = "C", TPOPVS_CAN, TPOPVS_OTR))
                    !tpomon = IIf(aRegistros(25) = "N", TPOMON_NAC, TPOMON_EXT)
                    !TpoTcb = IIf(aRegistros(26) = "C", TPOTCB_CPR, TPOTCB_VTA)
                    !ImpTCb = CDec(Format(Val(aRegistros(27)), FORMATO_NUM_2))
                    !ImpMN = CDec(Format(Val(aRegistros(28)), FORMATO_NUM_1))
                    !ImpME = CDec(Format(Val(aRegistros(29)), FORMATO_NUM_1))
                    !tpodoc = IIf(aRegistros(30) = "", Null, aRegistros(30))
                    !indfjo_det = INDPREGEN_INA
                    !IndGnr_RP = INDPREGEN_INA
                    !IndNCu = INDPREGEN_INA
                    !IndAnu = INDPREGEN_INA
                  End Select
                  !UsrCre = gsAbvUsr
                  !FyHCre = Now
                  .Update
                End With
                ' Actualizoel numero de comprobantes en los diarios  TC
              '  sSentencia = "UPDATE CoDro SET cpb" & gsMesAct & "='" & sComprobante & "' "
              '  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
              '  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
              '  sSentencia = sSentencia & "AND CodDro='" & sDiario & "'"
              '  pocnnMain.Execute sSentencia
                
                pgbProgreso(1).Value = IIf((Loc(nArchivo) * 128) > nNumRegistros, nNumRegistros, (Loc(nArchivo) * 128))
                DoEvents
              Loop
              porstTmp.Close
            End If
            Close #nArchivo
          End If
        Next nSecuencia
        ' Actualizo los tipos de cambio y equivalentes
        If chkVerificar.Value = vbChecked Then
          Select Case nContador
            Case 0    ' Registro de compras
              ' Actualizo el tipo de cambio
              sSentencia = "UPDATE tmpcocprdoc cpr, tgtcb tcb "
              sSentencia = sSentencia & "SET cpr.imptcb=tcb.imptcb_vta, cpr.indgen=" & INDPREGEN_ACT & " "
              sSentencia = sSentencia & "WHERE cpr.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND cpr.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND cpr.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND cpr.imptcb<=0 "
              sSentencia = sSentencia & "AND tcb.codemp=cpr.codemp "
              sSentencia = sSentencia & "AND tcb.fehtcb=cpr.feedoc"
              pocnnMain.Execute sSentencia, nNumRegistros
              ' Actualizo importes equivalentes de centro de costo
              If aActualiza(2) = INDPREGEN_ACT Then
                ' Actualizo importes equivalentes moneda nacional
                sSentencia = "UPDATE tmpcocprdoccco cco, tmpcocprdoc cpr "
                sSentencia = sSentencia & "SET cco.impcco_me = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cco.impcco_mn/cpr.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cco.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cco.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cpr.codemp=cco.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cco.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cco.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cco.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cco.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cco.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & ""
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Actualizo importes equivalentes moneda extranjera
                sSentencia = "UPDATE tmpcocprdoccco cco, tmpcocprdoc cpr "
                sSentencia = sSentencia & "SET cco.impcco_mn = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cco.impcco_me*cpr.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cco.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cco.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cpr.codemp=cco.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cco.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cco.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cco.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cco.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cco.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & ""
                pocnnMain.Execute sSentencia, nNumRegistros
              End If
            
              ' Actualizo importes equivalentes de cuentas
              If aActualiza(1) = INDPREGEN_ACT Then
                ' Sumatoria de centro de costos
                If aActualiza(2) = INDPREGEN_ACT Then
                  
                End If
                ' Actualizo importes equivalentes moneda nacional
                sSentencia = "UPDATE tmpcocprdoccta cta, tmpcocprdoc cpr "
                sSentencia = sSentencia & "SET cta.impcta_me = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cta.impcta_mn/cpr.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cpr.codemp=cta.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cta.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND cta.impcta_me=" & INDPREGEN_INA & ""
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Actualizo importes equivalentes moneda extranjera
                sSentencia = "UPDATE tmpcocprdoccta cta, tmpcocprdoc cpr "
                sSentencia = sSentencia & "SET cta.impcta_mn = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cta.impcta_me*cpr.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cpr.codemp=cta.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cta.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND cta.impcta_mn=" & INDPREGEN_INA & ""
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Elimino y genero el archivo sumatoria de cuentas
                sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & 12 & ")='#tmpsumacta_') DROP TABLE #tmpsumacta"
                pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsumacta", sSentencia)
                
                sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmpsumacta ", "")
                sSentencia = sSentencia & "SELECT codemp, pdoano, codaux, codtdc, serdoc, nrodoc,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 1 THEN impcta_mn ELSE 0.00 END), 2) AS impogr_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 2 THEN impcta_mn ELSE 0.00 END), 2) AS impogn_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 3 THEN impcta_mn ELSE 0.00 END), 2) AS impong_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 4 THEN impcta_mn ELSE 0.00 END), 2) AS impexo_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 5 THEN impcta_mn ELSE 0.00 END), 2) AS impigv_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 6 THEN impcta_mn ELSE 0.00 END), 2) AS impisc_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 7 THEN impcta_mn ELSE 0.00 END), 2) AS impoim_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 8 THEN impcta_mn ELSE 0.00 END), 2) AS impoi1_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 9 THEN impcta_mn ELSE 0.00 END), 2) AS impoi2_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 10 THEN impcta_mn ELSE 0.00 END), 2) AS impoi3_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 11 THEN impcta_mn ELSE 0.00 END), 2) AS imptot_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 11 THEN 0.00 ELSE impcta_mn END), 2) AS impsum_mn, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 1 THEN impcta_me ELSE 0.00 END), 2) AS impogr_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 2 THEN impcta_me ELSE 0.00 END), 2) AS impogn_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 3 THEN impcta_me ELSE 0.00 END), 2) AS impong_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 4 THEN impcta_me ELSE 0.00 END), 2) AS impexo_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 5 THEN impcta_me ELSE 0.00 END), 2) AS impigv_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 6 THEN impcta_me ELSE 0.00 END), 2) AS impisc_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 7 THEN impcta_me ELSE 0.00 END), 2) AS impoim_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 8 THEN impcta_me ELSE 0.00 END), 2) AS impoi1_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 9 THEN impcta_me ELSE 0.00 END), 2) AS impoi2_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 10 THEN impcta_me ELSE 0.00 END), 2) AS impoi3_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 11 THEN impcta_me ELSE 0.00 END), 2) AS imptot_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 11 THEN 0.00 ELSE impcta_me END), 2) AS impsum_me "
                sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & ps_Prefijo & "tmpsumacta ", "")
                sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcocprdoccta "
                sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "GROUP BY codemp, pdoano, codaux, codtdc, serdoc, nrodoc "
                sSentencia = sSentencia & "ORDER BY codaux, codtdc, serdoc, nrodoc"
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Sumo importe total cuenta diferente equivalente - moneda nacional
                sSentencia = "UPDATE tmpcocprdoccta cta, tmpcocprdoc cpr, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_me = Round(cta.impcta_me+(sum.impsum_me-sum.imptot_me), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_CPR & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND cpr.codemp=cta.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cta.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codaux=cta.codaux "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.impsum_me>sum.imptot_me"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Resto importe total cuenta diferente equivalente - moneda nacional
                sSentencia = "UPDATE tmpcocprdoccta cta, tmpcocprdoc cpr, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_me = Round(cta.impcta_me-(sum.imptot_me-sum.impsum_me), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_CPR & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND cpr.codemp=cta.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cta.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codaux=cta.codaux "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.imptot_me>sum.impsum_me"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Sumo importe total cuenta diferente equivalente - moneda extranjera
                sSentencia = "UPDATE tmpcocprdoccta cta, tmpcocprdoc cpr, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_mn = Round(cta.impcta_mn+(sum.impsum_mn-sum.imptot_mn), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_CPR & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND cpr.codemp=cta.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND cpr.codaux=cta.codaux "
                sSentencia = sSentencia & "AND cpr.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND cpr.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codaux=cta.codaux "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.impsum_mn>sum.imptot_mn"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Resto importe total cuenta diferente equivalente - moneda nacional
                sSentencia = "UPDATE tmpcocprdoccta cta, tmpcocprdoc cpr, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_mn = Round(cta.impcta_mn-(sum.imptot_mn-sum.impsum_mn), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_CPR & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND cpr.codemp=cta.codemp "
                sSentencia = sSentencia & "AND cpr.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND cpr.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND cpr.codaux=cta.codaux "
                sSentencia = sSentencia & "AND cpr.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND cpr.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND cpr.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND cpr.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codaux=cta.codaux "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.imptot_mn>sum.impsum_mn"
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Actualizo los importes equivalentes moneda nacional - sumatoria de cuentas
                sSentencia = "UPDATE tmpcocprdoc tmp, tmpsumacta sum SET "
                sSentencia = sSentencia & "tmp.impogr_me = sum.impogr_me, tmp.impogn_me = sum.impogn_me, tmp.impong_me = sum.impong_me, "
                sSentencia = sSentencia & "tmp.impexo_me = sum.impexo_me , tmp.impigv_me = sum.impigv_me, tmp.impisc_me = sum.impisc_me,"
                sSentencia = sSentencia & "tmp.impigv_ogr_me = (CASE WHEN sum.impogr_me>0 THEN sum.impigv_me ELSE 0 END), "
                sSentencia = sSentencia & "tmp.impigv_ogn_me = (CASE WHEN sum.impogn_me>0 THEN sum.impigv_me ELSE 0 END), "
                sSentencia = sSentencia & "tmp.impigv_ong_me = (CASE WHEN sum.impong_me>0 THEN sum.impigv_me ELSE 0 END), "
                sSentencia = sSentencia & "tmp.impoim_me = sum.impoim_me, tmp.impoi1_me = sum.impoi1_me, tmp.impoi2_me = sum.impoi2_me, "
                sSentencia = sSentencia & "tmp.impoi3_me = sum.impoi3_me, tmp.imptot_me = sum.impsum_me, tmp.indgen=" & INDPREGEN_INA & " "
                sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
                sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=tmp.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=tmp.pdoano "
                sSentencia = sSentencia & "AND sum.codaux=tmp.codaux "
                sSentencia = sSentencia & "AND sum.codtdc=tmp.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=tmp.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=tmp.nrodoc"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Actualizo los importes equivalentes moneda extranjera - sumatoria de cuentas
                sSentencia = "UPDATE tmpcocprdoc tmp, tmpsumacta sum SET "
                sSentencia = sSentencia & "tmp.impogr_mn = sum.impogr_mn, tmp.impogn_mn = sum.impogn_mn, tmp.impong_mn = sum.impong_mn, "
                sSentencia = sSentencia & "tmp.impexo_mn = sum.impexo_mn , tmp.impigv_mn = sum.impigv_mn, tmp.impisc_mn = sum.impisc_mn, "
                sSentencia = sSentencia & "tmp.impigv_ogr_mn = (CASE WHEN sum.impogr_mn>0 THEN sum.impigv_mn ELSE 0 END), "
                sSentencia = sSentencia & "tmp.impigv_ogn_mn = (CASE WHEN sum.impogn_mn>0 THEN sum.impigv_mn ELSE 0 END), "
                sSentencia = sSentencia & "tmp.impigv_ong_mn = (CASE WHEN sum.impong_mn>0 THEN sum.impigv_mn ELSE 0 END), "
                sSentencia = sSentencia & "tmp.impoim_mn = sum.impoim_mn, tmp.impoi1_mn = sum.impoi1_mn, tmp.impoi2_mn = sum.impoi2_mn, "
                sSentencia = sSentencia & "tmp.impoi3_mn = sum.impoi3_mn, tmp.imptot_mn = sum.impsum_mn, tmp.indgen=" & INDPREGEN_INA & " "
                sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
                sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=tmp.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=tmp.pdoano "
                sSentencia = sSentencia & "AND sum.codaux=tmp.codaux "
                sSentencia = sSentencia & "AND sum.codtdc=tmp.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=tmp.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=tmp.nrodoc"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Elimino y genero el archivo sumatoria de cuentas
                sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & 12 & ")='#tmpsumacta_') DROP TABLE #tmpsumacta"
                pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsumacta", sSentencia)
              End If
              ' Actualizo importes igv moneda nacional
              sSentencia = "UPDATE tmpcocprdoc tmp "
              sSentencia = sSentencia & "SET "
              sSentencia = sSentencia & "tmp.impigv_ogr_mn = (CASE WHEN tmp.impogr_mn>0 THEN tmp.impigv_mn ELSE 0 END), "
              sSentencia = sSentencia & "tmp.impigv_ogn_mn = (CASE WHEN tmp.impogn_mn>0 THEN tmp.impigv_mn ELSE 0 END), "
              sSentencia = sSentencia & "tmp.impigv_ong_mn = (CASE WHEN tmp.impong_mn>0 THEN tmp.impigv_mn ELSE 0 END) "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_NAC & "' "
              sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & ""
              pocnnMain.Execute sSentencia, nNumRegistros
              ' Actualizo importes igv moneda extranjera
              sSentencia = "UPDATE tmpcocprdoc tmp "
              sSentencia = sSentencia & "SET "
              sSentencia = sSentencia & "tmp.impigv_ogr_me = (CASE WHEN tmp.impogr_me>0 THEN tmp.impigv_me ELSE 0 END), "
              sSentencia = sSentencia & "tmp.impigv_ogn_me = (CASE WHEN tmp.impogn_me>0 THEN tmp.impigv_me ELSE 0 END), "
              sSentencia = sSentencia & "tmp.impigv_ong_me = (CASE WHEN tmp.impong_me>0 THEN tmp.impigv_me ELSE 0 END) "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_EXT & "' "
              sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & ""
              pocnnMain.Execute sSentencia, nNumRegistros
              
              ' Actualizo los importes equivalentes moneda nacional
              sSentencia = "UPDATE tmpcocprdoc tmp "
              sSentencia = sSentencia & "SET tmp.impogr_me = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(tmp.impogr_mn/tmp.imptcb, 2), 0), "
              sSentencia = sSentencia & "tmp.impogn_me = Round(tmp.impogn_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impong_me = Round(tmp.impong_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impexo_me = Round(tmp.impexo_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_me = Round(tmp.impigv_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_ogr_me = Round(tmp.impigv_ogr_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_ogn_me = Round(tmp.impigv_ogn_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_ong_me = Round(tmp.impigv_ong_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impisc_me = Round(tmp.impisc_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoim_me = Round(tmp.impoim_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoi1_me = Round(tmp.impoi1_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoi2_me = Round(tmp.impoi2_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoi3_me = Round(tmp.impoi3_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.imptot_me = Round(tmp.imptot_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.indgen=" & INDPREGEN_INA & " "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_NAC & "' "
              sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
              pocnnMain.Execute sSentencia, nNumRegistros
              
              ' Actualizo los importes equivalentes moneda extranjera
              sSentencia = "UPDATE tmpcocprdoc tmp "
              sSentencia = sSentencia & "SET tmp.impogr_mn = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(tmp.impogr_me*tmp.imptcb, 2), 0), "
              sSentencia = sSentencia & "tmp.impogn_mn = Round(tmp.impogn_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impong_mn = Round(tmp.impong_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impexo_mn = Round(tmp.impexo_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_mn = Round(tmp.impigv_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_ogr_mn = Round(tmp.impigv_ogr_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_ogn_mn = Round(tmp.impigv_ogn_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_ong_mn = Round(tmp.impigv_ong_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impisc_mn = Round(tmp.impisc_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoim_mn = Round(tmp.impoim_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoi1_mn = Round(tmp.impoi1_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoi2_mn = Round(tmp.impoi2_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoi3_mn = Round(tmp.impoi3_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.imptot_mn = Round(tmp.imptot_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.indgen=" & INDPREGEN_INA & " "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_EXT & "' "
              sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
              pocnnMain.Execute sSentencia, nNumRegistros
            Case 1    ' Registro de ventas
              ' Actualizo el tipo de cambio
              sSentencia = "UPDATE tmpcovtadoc vta, tgtcb tcb "
              sSentencia = sSentencia & "SET vta.imptcb=tcb.imptcb_vta, vta.indgen=" & INDPREGEN_ACT & " "
              sSentencia = sSentencia & "WHERE vta.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND vta.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND vta.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND vta.imptcb<=0 "
              sSentencia = sSentencia & "AND tcb.codemp=vta.codemp "
              sSentencia = sSentencia & "AND tcb.fehtcb=vta.feedoc"
              pocnnMain.Execute sSentencia, nNumRegistros
              
              ' Actualizo importes equivalentes de centro de costo
              If aActualiza(2) = INDPREGEN_ACT Then
                ' Actualizo importes equivalentes moneda nacional
                sSentencia = "UPDATE tmpcovtadoccco cco, tmpcovtadoc vta "
                sSentencia = sSentencia & "SET cco.impcco_me = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cco.impcco_mn/vta.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cco.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cco.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND vta.codemp=cco.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cco.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cco.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cco.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cco.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & ""
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Actualizo importes equivalentes moneda extranjera
                sSentencia = "UPDATE tmpcovtadoccco cco, tmpcovtadoc vta "
                sSentencia = sSentencia & "SET cco.impcco_mn = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cco.impcco_me*vta.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cco.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cco.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND vta.codemp=cco.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cco.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cco.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cco.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cco.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & ""
                pocnnMain.Execute sSentencia, nNumRegistros
              End If
              
              ' Actualizo importes equivalentes de cuentas
              If aActualiza(1) = INDPREGEN_ACT Then
                ' Sumatoria de centro de costos
                If aActualiza(2) = INDPREGEN_ACT Then
                  
                End If
                ' Actualizo importes equivalentes moneda nacional
                sSentencia = "UPDATE tmpcovtadoccta cta, tmpcovtadoc vta "
                sSentencia = sSentencia & "SET cta.impcta_me = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cta.impcta_mn/vta.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND vta.codemp=cta.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND cta.impcta_me=" & INDPREGEN_INA & ""
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Actualizo importes equivalentes moneda extranjera
                sSentencia = "UPDATE tmpcovtadoccta cta, tmpcovtadoc vta "
                sSentencia = sSentencia & "SET cta.impcta_mn = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(cta.impcta_me*vta.imptcb, 2), 0) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND vta.codemp=cta.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND cta.impcta_mn=" & INDPREGEN_INA & ""
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Elimino y genero el archivo sumatoria de cuentas
                sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & 12 & ")='#tmpsumacta_') DROP TABLE #tmpsumacta"
                pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsumacta", sSentencia)
                
                sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmpsumacta ", "")
                sSentencia = sSentencia & "SELECT codemp, pdoano, codtdc, serdoc, nrodoc,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 1 THEN impcta_mn ELSE 0.00 END), 2) AS impogr_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 2 THEN impcta_mn ELSE 0.00 END), 2) AS impexp_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 3 THEN impcta_mn ELSE 0.00 END), 2) AS impexo_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 4 THEN impcta_mn ELSE 0.00 END), 2) AS impigv_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 5 THEN impcta_mn ELSE 0.00 END), 2) AS impisc_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 6 THEN impcta_mn ELSE 0.00 END), 2) AS impoim_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 7 THEN impcta_mn ELSE 0.00 END), 2) AS imptot_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 7 THEN 0.00 ELSE impcta_mn END), 2) AS impsum_mn,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 1 THEN impcta_me ELSE 0.00 END), 2) AS impogr_me,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 2 THEN impcta_me ELSE 0.00 END), 2) AS impexp_me,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 3 THEN impcta_me ELSE 0.00 END), 2) AS impexo_me,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 4 THEN impcta_me ELSE 0.00 END), 2) AS impigv_me,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 5 THEN impcta_me ELSE 0.00 END), 2) AS impisc_me,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 6 THEN impcta_me ELSE 0.00 END), 2) AS impoim_me,"
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 7 THEN impcta_me ELSE 0.00 END), 2) AS imptot_me, "
                sSentencia = sSentencia & "ROUND(SUM(CASE tpocnc WHEN 7 THEN 0.00 ELSE impcta_me END), 2) AS impsum_me "
                sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & ps_Prefijo & "tmpsumacta ", "")
                sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcovtadoccta "
                sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "GROUP BY codemp, pdoano, codtdc, serdoc, nrodoc "
                sSentencia = sSentencia & "ORDER BY codtdc, serdoc, nrodoc"
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Sumo importe total cuenta diferente equivalente - moneda nacional
                sSentencia = "UPDATE tmpcovtadoccta cta, tmpcovtadoc vta, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_me = Round(cta.impcta_me+(sum.impsum_me-sum.imptot_me), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_VTA & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND vta.codemp=cta.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.impsum_me>sum.imptot_me"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Resto importe total cuenta diferente equivalente - moneda nacional
                sSentencia = "UPDATE tmpcovtadoccta cta, tmpcovtadoc vta, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_me = Round(cta.impcta_me-(sum.imptot_me-sum.impsum_me), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_VTA & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND vta.codemp=cta.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.imptot_me>sum.impsum_me"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Sumo importe total cuenta diferente equivalente - moneda extranjera
                sSentencia = "UPDATE tmpcovtadoccta cta, tmpcovtadoc vta, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_mn = Round(cta.impcta_mn+(sum.impsum_mn-sum.imptot_mn), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_VTA & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND vta.codemp=cta.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.impsum_mn>sum.imptot_mn"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Resto importe total cuenta diferente equivalente - moneda nacional
                sSentencia = "UPDATE tmpcovtadoccta cta, tmpcovtadoc vta, tmpsumacta sum "
                sSentencia = sSentencia & "SET cta.impcta_mn = Round(cta.impcta_mn-(sum.imptot_mn-sum.impsum_mn), 2) "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND cta.tpocnc=" & TPOCNC_TOT_VTA & " "
                sSentencia = sSentencia & "AND cta.orden='01' "
                sSentencia = sSentencia & "AND vta.codemp=cta.codemp "
                sSentencia = sSentencia & "AND vta.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND vta.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND vta.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND vta.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND vta.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND vta.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=cta.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND sum.codtdc=cta.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=cta.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=cta.nrodoc "
                sSentencia = sSentencia & "AND sum.imptot_mn>sum.impsum_mn"
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Actualizo los importes equivalentes moneda nacional - sumatoria de cuentas
                sSentencia = "UPDATE tmpcovtadoc tmp, tmpsumacta sum "
                sSentencia = sSentencia & "SET tmp.impogr_me = sum.impogr_me, tmp.impexp_me = sum.impexp_me, "
                sSentencia = sSentencia & "tmp.impexo_me = sum.impexo_me , tmp.impigv_me = sum.impigv_me, "
                sSentencia = sSentencia & "tmp.impisc_me = sum.impisc_me, tmp.impoim_me = sum.impoim_me, "
                sSentencia = sSentencia & "tmp.imptot_me = sum.impsum_me, tmp.indgen=" & INDPREGEN_INA & " "
                sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
                sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_NAC & "' "
                sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=tmp.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=tmp.pdoano "
                sSentencia = sSentencia & "AND sum.codtdc=tmp.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=tmp.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=tmp.nrodoc"
                pocnnMain.Execute sSentencia, nNumRegistros
                ' Actualizo los importes equivalentes moneda extranjera - sumatoria de cuentas
                sSentencia = "UPDATE tmpcovtadoc tmp, tmpsumacta sum "
                sSentencia = sSentencia & "SET tmp.impogr_mn = sum.impogr_mn, tmp.impexp_mn = sum.impexp_mn, "
                sSentencia = sSentencia & "tmp.impexo_mn = sum.impexo_mn , tmp.impigv_mn = sum.impigv_mn, "
                sSentencia = sSentencia & "tmp.impisc_mn = sum.impisc_mn, tmp.impoim_mn = sum.impoim_mn, "
                sSentencia = sSentencia & "tmp.imptot_mn = sum.impsum_mn, tmp.indgen=" & INDPREGEN_INA & " "
                sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
                sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_EXT & "' "
                sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
                sSentencia = sSentencia & "AND sum.codemp=tmp.codemp "
                sSentencia = sSentencia & "AND sum.pdoano=tmp.pdoano "
                sSentencia = sSentencia & "AND sum.codtdc=tmp.codtdc "
                sSentencia = sSentencia & "AND sum.serdoc=tmp.serdoc "
                sSentencia = sSentencia & "AND sum.nrodoc=tmp.nrodoc"
                pocnnMain.Execute sSentencia, nNumRegistros
                
                ' Elimino y genero el archivo sumatoria de cuentas
                sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & 12 & ")='#tmpsumacta_') DROP TABLE #tmpsumacta"
                pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpsumacta", sSentencia)
              End If
            
              ' Actualizo los importes equivalentes moneda nacional
              sSentencia = "UPDATE tmpcovtadoc tmp "
              sSentencia = sSentencia & "SET tmp.impogr_me = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(tmp.impogr_mn/tmp.imptcb, 2), 0), "
              sSentencia = sSentencia & "tmp.impexp_me = Round(tmp.impexp_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impexo_me = Round(tmp.impexo_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_me = Round(tmp.impigv_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impisc_me = Round(tmp.impisc_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoim_me = Round(tmp.impoim_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.imptot_me = Round(tmp.imptot_mn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.indgen=" & INDPREGEN_INA & " "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_NAC & "' "
              sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & ""
              pocnnMain.Execute sSentencia, nNumRegistros
              ' Actualizo los importes equivalentes moneda extranjera
              sSentencia = "UPDATE tmpcovtadoc tmp "
              sSentencia = sSentencia & "SET tmp.impogr_mn = " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(Round(tmp.impogr_me*tmp.imptcb, 2), 0), "
              sSentencia = sSentencia & "tmp.impexp_mn = Round(tmp.impexp_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impexo_mn = Round(tmp.impexo_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impigv_mn = Round(tmp.impigv_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impisc_mn = Round(tmp.impisc_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.impoim_mn = Round(tmp.impoim_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.imptot_mn = Round(tmp.imptot_me*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.indgen=" & INDPREGEN_INA & " "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs='" & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_EXT & "' "
              sSentencia = sSentencia & "AND tmp.indgen=" & INDPREGEN_ACT & " "
              pocnnMain.Execute sSentencia, nNumRegistros
            Case 2    ' Registro de honorarios
            Case 3    ' Registro de diarios
              ' Actualizo el tipo de cambio de provisiones
              sSentencia = "UPDATE tmpcocpbdet dro, tgtcb tcb "
              sSentencia = sSentencia & "SET dro.imptcb=(CASE dro.tpotcb WHEN '" & TPOTCB_CPR & "' THEN tcb.imptcb_cpr ELSE tcb.imptcb_vta END), dro.indanu=" & INDANU_VER & " "
              sSentencia = sSentencia & "WHERE dro.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND dro.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND dro.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
              sSentencia = sSentencia & "AND dro.tpopvs='" & TPOPVS_PVS & "' "
              sSentencia = sSentencia & "AND dro.imptcb<=0 "
              sSentencia = sSentencia & "AND tcb.codemp=dro.codemp "
              sSentencia = sSentencia & "AND tcb.fehtcb=dro.feedoc"
              pocnnMain.Execute sSentencia, nNumRegistros
              ' Actualizo el tipo de cambio de otras transacciones
              sSentencia = "UPDATE tmpcocpbdet dro, tgtcb tcb "
              sSentencia = sSentencia & "SET dro.imptcb=(CASE dro.tpotcb WHEN '" & TPOTCB_CPR & "' THEN tcb.imptcb_cpr ELSE tcb.imptcb_vta END), dro.indanu=" & INDANU_VER & " "
              sSentencia = sSentencia & "WHERE dro.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND dro.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND dro.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
              sSentencia = sSentencia & "AND dro.tpopvs<>'" & TPOPVS_PVS & "' "
              sSentencia = sSentencia & "AND dro.imptcb<=0 "
              sSentencia = sSentencia & "AND tcb.codemp=dro.codemp "
              sSentencia = sSentencia & "AND tcb.fehtcb=dro.fehope"
              pocnnMain.Execute sSentencia, nNumRegistros
              
              ' Actualizo los importes equivalentes moneda nacional
              sSentencia = "UPDATE tmpcocpbdet tmp "
              sSentencia = sSentencia & "SET tmp.impme = Round(tmp.impmn/tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.indanu=" & INDANU_FAL & " "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_NAC & "' "
              sSentencia = sSentencia & "AND tmp.indanu=" & INDANU_VER & ""
              pocnnMain.Execute sSentencia, nNumRegistros
              ' Actualizo los importes equivalentes moneda extranjera
              sSentencia = "UPDATE tmpcocpbdet tmp "
              sSentencia = sSentencia & "SET tmp.impmn = Round(tmp.impme*tmp.imptcb, 2), "
              sSentencia = sSentencia & "tmp.indanu=" & INDANU_FAL & " "
              sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND tmp.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
              sSentencia = sSentencia & "AND tmp.tpomon='" & TPOMON_EXT & "' "
              sSentencia = sSentencia & "AND tmp.indanu=" & INDANU_VER & ""
              pocnnMain.Execute sSentencia, nNumRegistros
          End Select
        End If
        ' Cuentas registradas
        Select Case nContador
         Case 0    ' Registro de compras
          ' Actualizo indicador de cuentas registradas
          For nSecuencia = 1 To TPOCNC_TOT_CPR
            sSentencia = "UPDATE tmpcocprdoc cpr, tmpcocprdoccta cta "
            sSentencia = sSentencia & "SET cpr.indcta_" & Choose(nSecuencia, "ogr", "ogn", "ong", "exo", "isc", "igv", "oim", "oi1", "oi2", "oi3", "tot") & "=" & INDPREGEN_ACT & " "
            sSentencia = sSentencia & "WHERE cpr.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND cpr.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND cta.codemp=cpr.codemp "
            sSentencia = sSentencia & "AND cta.pdoano=cpr.pdoano "
            sSentencia = sSentencia & "AND cta.codaux=cpr.codaux "
            sSentencia = sSentencia & "AND cta.codtdc=cpr.codtdc "
            sSentencia = sSentencia & "AND cta.serdoc=cpr.serdoc "
            sSentencia = sSentencia & "AND cta.nrodoc=cpr.nrodoc "
            sSentencia = sSentencia & "AND cta.tpocnc=" & nSecuencia & " "
            sSentencia = sSentencia & "AND cta.orden>='01'"
            pocnnMain.Execute sSentencia, nNumRegistros
          Next nSecuencia
         Case 1    ' Registro de ventas
          If aActualiza(1) = INDPREGEN_ACT Then
            ' Actualizo indicador de cuentas registradas
            For nSecuencia = 1 To TPOCNC_TOT_VTA
              sSentencia = "UPDATE tmpcovtadoc vta, tmpcovtadoccta cta "
              sSentencia = sSentencia & "SET vta.indcta_" & Choose(nSecuencia, "ogr", "exp", "exo", "igv", "isc", "oim", "tot") & "=" & INDPREGEN_ACT & " "
              sSentencia = sSentencia & "WHERE vta.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND vta.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND cta.codemp=vta.codemp "
              sSentencia = sSentencia & "AND cta.pdoano=vta.pdoano "
              sSentencia = sSentencia & "AND cta.codtdc=vta.codtdc "
              sSentencia = sSentencia & "AND cta.serdoc=vta.serdoc "
              sSentencia = sSentencia & "AND cta.nrodoc=vta.nrodoc "
              sSentencia = sSentencia & "AND cta.tpocnc=" & nSecuencia & " "
              sSentencia = sSentencia & "AND cta.orden>='01'"
              pocnnMain.Execute sSentencia, nNumRegistros
            Next nSecuencia
          End If
         Case 2    ' Registro de honorarios
        End Select
      End If
    End If
  Next nContador
  Set porstTmp = Nothing

End Sub

Private Sub ppImporta_Tablas()
    Static sSentencia As String, sTabla As String
    Static sArchivo As String, sMilinea As String
    Static nContador As Integer, nArchivo As Integer
    Static nColumnas As Integer
    Static nRegistro As Double, nNumRegistros As Double
    Static aRegistros()
    Static porstTmp As ADODB.Recordset

    ' Seteo el recordset temporal para la grabacion
    Set porstTmp = New ADODB.Recordset
    ' Importo las tablas de acuerdo a la selección
    nArchivo = FreeFile
    For nContador = 0 To chkImporTabla.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporTabla(nContador).Value Then
        ' Abro Archivo de Texto
        sArchivo = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & Choose(nContador + 1, "dro", "pct", "aux", "tdc", "cco", "tpc") & ".txt"
        nColumnas = Choose(nContador + 1, "3", "30", "11", "5", "3", "3")
        ' Desactivo la opcion si no existe archivo
        chkImporTabla(nContador).Value = vbUnchecked
        If Dir$(sArchivo, vbNormal) <> "" Then
          ' Activo la opcion si existe archivo
          chkImporTabla(nContador).Value = vbChecked
          Open sArchivo For Input As #nArchivo
          nNumRegistros = gfRedond(LOF(nArchivo), 0)
          If nNumRegistros > 0 Then
            pgbProgreso(1).Max = nNumRegistros
            pgbProgreso(1).Value = pgbProgreso(1).Min
            ' Genero el archivo y abro el recordset temporal
            sTabla = Choose(nContador + 1, "CoDro", "CoCta", "TgAux", "TgTDc", "CoCCo", "TgTcb")
            sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & Len(sTabla) + 5 & ")='#tmp" & Len(sTabla) & "_') DROP TABLE #tmp" & sTabla
            pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp" & sTabla, sSentencia)
            
            sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmp" & sTabla & " ", "")
             
            sSentencia = sSentencia & "SELECT * "
            If nContador = 2 Then
              sSentencia = sSentencia & ", space(20) AS NomAux, space(20) AS ApePatAux, space(20) AS ApeMatAux "
            End If
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & ps_Prefijo & "tmp" & sTabla & " ", "")
            sSentencia = sSentencia & "FROM " & sTabla & " "
            sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND UsrCre='tmpusrsys'"
            pocnnMain.Execute sSentencia
            With porstTmp
              If .State = adStateOpen Then .Close
              .ActiveConnection = pocnnMain
              .Source = "SELECT * FROM " & ps_Prefijo & "tmp" & sTabla
              .CursorType = adOpenDynamic
              .LockType = adLockOptimistic
              .Open
            End With
            ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
            ReDim aRegistros(nColumnas)
            lblProgreso(1).Caption = Choose(gsIdioma, "Importando Archivo: ", "importing File: ") & Trim(chkImporTabla(nContador).Caption) & " (" & gsRUCEmp & Choose(nContador + 1, "dro", "pct", "aux", "tdc", "cco", "tpc") & ")"
            Do While Not EOF(nArchivo)
              Line Input #nArchivo, sMilinea
              nRegistro = nRegistro + 1
              pRegistro_Texto sMilinea, nColumnas, aRegistros
              With porstTmp
                .AddNew
                !codemp = gsCodEmp
                Select Case nContador
                 Case 0
                  !pdoano = gsAnoAct
                  !coddro = aRegistros(2)
                  !DetDro = aRegistros(3)
                 Case 1
                 
                  '!pdoano = gsAnoAct
                  '!codcta = aRegistros(2)
                  '!detcta = Left(aRegistros(3), 60)
                  '!TpoCTA = IIf(aRegistros(4) = "D", TPOCTA_TRA, TPOCTA_TIT)
                  '!NatCta = IIf(aRegistros(5) = "D", NATCTA_DEU, NATCTA_ACR)
                  '!TpoSdo = IIf(aRegistros(6) = "I", TPOSDO_INV, IIf(aRegistros(6) = "R", TPOSDO_RES, IIf(aRegistros(6) = "F", TPOSDO_FUN, IIf(aRegistros(6) = "N", TPOSDO_NAT, TPOSDO_AMB))))
                  '!TpoAnl = IIf(aRegistros(7) = "C", TPOANL_CTA, IIf(aRegistros(7) = "A", TPOANL_AUX, TPOANL_DOC))
                  '!CodCta_Dst_Deb = aRegistros(8)
                  '!CodCta_Dst_Hab = aRegistros(9)
                  '!tpomon = IIf(aRegistros(10) = "N", TPOMON_NAC, TPOMON_EXT)
                  '!TpoTcb = IIf(aRegistros(11) = "C", TPOTCB_CPR, TPOTCB_VTA)
                  '!tpoajd = IIf(aRegistros(12) = "S", INDAJD_ACT, INDAJD_INA)
                  '!CodCta_AjD_Deb = aRegistros(13)
                  '!CodCta_AjD_Hab = aRegistros(14)
                  '!IndAjD = IIf(aRegistros(15) = "S", INDAJD_ACT, INDAJD_INA)
                  '!IndCCo = IIf(aRegistros(16) = "S", INDCCO_ACT, INDCCO_INA)
                  '!IndDoc = IIf(aRegistros(17) = "S", INDDOC_ACT, INDDOC_INA)
                  '!IndMoe = IIf(aRegistros(18) = "S", INDMOE_ACT, INDMOE_INA)
                  '!IndPsp = IIf(aRegistros(19) = "S", INDPSP_ACT, INDPSP_INA)
                  '!Indfjo = INDFJO_INA
                  '!estcta = ESTCTA_ACT
                  
                  !pdoano = gsAnoAct
                  !codcta = aRegistros(2)
                  !detcta = Left(aRegistros(3), 60)
                  !DetCtax = "xxxxxxxx"
                  !TpoCTA = IIf(aRegistros(5) = "D", TPOCTA_TRA, TPOCTA_TIT)
                  !NatCta = IIf(aRegistros(6) = "D", NATCTA_DEU, NATCTA_ACR)
                  !TpoSdo = IIf(aRegistros(7) = "I", TPOSDO_INV, IIf(aRegistros(7) = "R", TPOSDO_RES, IIf(aRegistros(7) = "F", TPOSDO_FUN, IIf(aRegistros(7) = "N", TPOSDO_NAT, TPOSDO_AMB))))
                  !TpoAnl = IIf(aRegistros(8) = "C", TPOANL_CTA, IIf(aRegistros(8) = "A", TPOANL_AUX, TPOANL_DOC))
                  !codcta_dst_deb = IIf(Trim$(aRegistros(9)) <> "", aRegistros(9), Null)
                  !codcta_dst_hab = IIf(Trim$(aRegistros(10)) <> "", aRegistros(10), Null)
                  !codcco_dst_deb = IIf(Trim$(aRegistros(11)) <> "", aRegistros(11), Null)
                  !codcco_dst_hab = IIf(Trim$(aRegistros(12)) <> "", aRegistros(12), Null)
                  !tpomon = IIf(aRegistros(13) = "N", TPOMON_NAC, TPOMON_EXT)
                  !TpoTcb = IIf(aRegistros(14) = "C", TPOTCB_CPR, TPOTCB_VTA)
                  !tpoajd = IIf(aRegistros(15) = "S", INDAJD_ACT, INDAJD_INA)
                  !CodCta_AjD_Deb = IIf(Trim$(aRegistros(16)) <> "", aRegistros(16), Null)
                  !CodCta_AjD_Hab = IIf(Trim$(aRegistros(17)) <> "", aRegistros(17), Null)
                  !CodCCo_AjD_Deb = IIf(Trim$(aRegistros(18)) <> "", aRegistros(18), Null)
                  !CodCCo_AjD_Hab = IIf(Trim$(aRegistros(19)) <> "", aRegistros(19), Null)
                  !IndAjD = IIf(aRegistros(20) = "S", INDAJD_ACT, INDAJD_INA)
                  !codcta_crr_deu = IIf(Trim$(aRegistros(21)) <> "", aRegistros(21), Null)
                  !codcta_crr_acr = IIf(Trim$(aRegistros(22)) <> "", aRegistros(22), Null)
'ini 2015-05-10 deshabilitado para
'pasar plan de cuentas emco
'****************************
'lo puse en su estado original
'                  !codcco_def = Null
'                  !indcco = IIf(aRegistros(23) = "S", INDCCO_ACT, INDCCO_INA)
'                  !IndDoc = IIf(aRegistros(24) = "S", INDDOC_ACT, INDDOC_INA)
'                  !IndMoe = IIf(aRegistros(25) = "S", INDMOE_ACT, INDMOE_INA)
'                  !IndPsp = IIf(aRegistros(26) = "S", INDPSP_ACT, INDPSP_INA)
'                  !IndFjo = INDFJO_INA
'                  '!Codbco = aRegistros(28)
'                  !codbco = IIf(Trim$(aRegistros(28)) <> "", aRegistros(28), Null)
'                  '2015-05-10 error null cambiado !EstCta = ESTCTA_ACT
'                  !EstCta = IIf(Trim$(aRegistros(29)) <> "", aRegistros(29), Null)
'****************************
'este es el original
                  !codcco_def = IIf(Trim$(aRegistros(23)) <> "", aRegistros(23), Null)
                  !indcco = IIf(aRegistros(24) = "S", INDCCO_ACT, INDCCO_INA)
                  !IndDoc = IIf(aRegistros(25) = "S", INDDOC_ACT, INDDOC_INA)
                  !IndMoe = IIf(aRegistros(26) = "S", INDMOE_ACT, INDMOE_INA)
                  !IndPsp = IIf(aRegistros(27) = "S", INDPSP_ACT, INDPSP_INA)
                  !IndFjo = INDFJO_INA
                  '!Codbco = aRegistros(28)
                  !codbco = IIf(Trim$(aRegistros(29)) <> "", aRegistros(29), Null)
                  '2015-05-10 error null cambiado !EstCta = ESTCTA_ACT
                  !EstCta = IIf(Trim$(aRegistros(30)) <> "", aRegistros(30), Null)
'2015-05-10 deshabilitado para
'pasar plan de cuentas emco
                 Case 2
                  !codaux = aRegistros(1)
                  !razAux = Left(IIf(Trim(aRegistros(3)) <> "", Trim$(aRegistros(4)) & " " & Trim$(aRegistros(5)) & "," & Trim$(aRegistros(3)), Trim$(aRegistros(2))), 60)
                  !NomAux = Left(Trim(aRegistros(3)), 20)
                  !ApePatAux = Left(Trim(aRegistros(4)), 20)
                  !ApeMatAux = Left(Trim(aRegistros(5)), 20)
                  !rucaux = Left(Trim(aRegistros(6)), 11)
                  '2014-05-20 para corregir errot txt sale en tipo numercion correlativa 1,2,3...n
                  '!TpoDci = Left(Trim(aRegistros(11)), 2)
                  If gsCodEmpCompass = CODEMP_COMPASS Then
                    !TpoDci = "06"
                  Else
                    !TpoDci = Left(Trim(aRegistros(11)), 2)
                  End If
                  !DirAux = Left(Trim(aRegistros(7)), 80)
                  !IndCli = IIf(aRegistros(8) = "S", INDAUX_CLI_ACT, INDAUX_CLI_INA)
                  !IndPrv = IIf(aRegistros(9) = "S", INDAUX_PRV_ACT, INDAUX_PRV_INA)
                  !IndOtr = IIf(aRegistros(10) = "S", INDAUX_OTR_ACT, INDAUX_OTR_INA)
                  !TpoPer = IIf(Trim$(aRegistros(3)) <> "", TPOPER_NAT, TPOPER_JUR)
                  !EstAux = ESTAUX_ACT
                 Case 3
                  !codtdc = aRegistros(1)
                  !dettdc = aRegistros(2)
                  !AbvTDc = aRegistros(3)
                  !SgnTDc = IIf(aRegistros(4) = "+", SGNTDC_POS, SGNTDC_NEG)
                  '!SgnTDc = IIf(aRegistros(5) = "", "0", aRegistros(5))
                  !forimp = "0"
                 Case 4
                  !pdoano = gsAnoAct
                  !codcco = aRegistros(2)
                  !detcco = aRegistros(3)
                  !EstCCo = ESTCCO_ACT
                 Case 5
                  !FehTCb = Format(aRegistros(1), "dd/mm/yyyy")
                  !ImpTCb_Cpr = Format(CDbl(aRegistros(2)), FORMATO_NUM_2)
                  !ImpTCb_Vta = Format(CDbl(aRegistros(3)), FORMATO_NUM_2)
                End Select
                !UsrCre = gsAbvUsr
                !FyHCre = Now
                .Update
              End With
              pgbProgreso(1).Value = IIf((Loc(nArchivo) * 128) > nNumRegistros, nNumRegistros, (Loc(nArchivo) * 128))
              DoEvents
            Loop
            porstTmp.Close
          End If
          Close #nArchivo
        End If
      End If
    Next nContador
    Set porstTmp = Nothing

End Sub

Private Sub ppInicializa_Tablas()
  Dim sTabla As String, sColumnas As String, sWhere As String
  Dim nRegistro As Long, nNumRegistros As Long
  Dim sSentencia As String
  Dim nContador As Integer

  pgbProgreso(0).Max = chkTransTabla.Count
  pgbProgreso(0).Value = pgbProgreso(0).Min
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkTransTabla.Count - 1
    ' Verifico que se haya seleccionado
    If (chkTransTabla(nContador).Value = vbChecked) Then
      ' Obtengo el archivo de texto
      sColumnas = Choose(nContador + 1, "a.coddro, a.detdro, a.detdrox", "a.codbco, a.forimp, a.detbco, a.detbcox, a.codent, a.ctactemn, a.ctacteme", "a.codcta, a.detcta, a.detctax, a.tpocta, a.natcta, a.tposdo, a.tpoanl, a.codcta_dst_deb, a.codcta_dst_hab, a.codcco_dst_deb, a.codcco_dst_hab, a.tpomon, a.tpotcb, a.tpoajd, a.codcta_ajd_deb, a.codcta_ajd_hab, a.codcco_ajd_deb, a.codcco_ajd_hab, a.indajd, a.codcta_crr_deu, a.codcta_crr_acr, a.codcco_def, a.indcco, a.inddoc, a.indmoe, a.indpsp, a.indfjo, a.codbco, a.estcta", "a.codaux, a.razaux, a.rucaux, a.tpodci, a.diraux, a.rubro, a.email, a.indcli, a.indprv, a.indotr, a.tpoper, a.estaux", "a.codtdc, a.dettdc, a.dettdcx, a.abvtdc, a.sgntdc, a.forimp", "a.codcco, a.detcco, a.detccox, a.indpdocpr, a.estcco", "a.fehtcb, a.imptcb_cpr, a.imptcb_vta, a.imptcb_bco_cpr, a.imptcb_bco_vta", "a.codefi, a.detefi, a.detefix, a.coddpe, a.indcnv", "a.codasi, a.detasi, a.detasix, a.tpoasi", "a.codefe, a.detefe, a.detefex, a.tpoefe")
      sTabla = Choose(nContador + 1, "codro", "cobco", "cocta", "tgaux", "tgtdc", "cocco", "tgtcb", "coefi", "coasitipo", "coefe")
      sWhere = Choose(nContador + 1, ", pdoano", "", ", pdoano", "", "", ", pdoano", "", ", pdoano", ", pdoano", ", pdoano")
        
      ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
      lblProgreso(1).Caption = Choose(gsIdioma, "Exportando Información: ", "Exporting Information: ") & Trim(chkTransTabla(nContador).Caption)
      sSentencia = "INSERT INTO " & sTabla & " (" & Replace(sColumnas, "a.", "") & sWhere & ", codemp, usrcre, fyhcre) "
      Select Case nContador
       Case 0     ' diario
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.coddro=a.coddro) "
        sSentencia = sSentencia & "ORDER BY a.coddro"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' libro auxiliar
        sSentencia = "INSERT INTO colib (codlib, deslib, estadolib, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT a.codlib, a.deslib, a.estadolib, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM colib a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM colib b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.codlib=a.codlib) "
        sSentencia = sSentencia & "ORDER BY a.codlib"
       Case 1     ' entidad banco
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.codbco=a.codbco) "
        sSentencia = sSentencia & "ORDER BY a.codbco"
       Case 2     ' plan de cuentas
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codcta=a.codcta) "
        sSentencia = sSentencia & "ORDER BY a.codcta"
       Case 3     ' auxiliar
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.codaux=a.codaux) "
        sSentencia = sSentencia & "ORDER BY a.codaux"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' personas naturales
        sSentencia = "INSERT INTO tgauxnat (codaux, nomaux, apepataux, apemataux, codtdi, numdci, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT codaux, nomaux, apepataux, apemataux, codtdi, numdci, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM tgauxnat a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM tgauxnat b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.codaux=a.codaux) "
        sSentencia = sSentencia & "ORDER BY a.codaux"
       Case 4     ' tipo documento
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.codtdc=a.codtdc) "
        sSentencia = sSentencia & "ORDER BY a.codtdc"
       Case 5     ' centro costo
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codcco=a.codcco) "
        sSentencia = sSentencia & "ORDER BY a.codcco"
       Case 6     ' tipo de cambio
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND DATE_FORMAT(a.fehtcb, '%Y') >='" & (CLng(Right(cboEjercicio.Text, 4)) - 1) & "' "
        sSentencia = sSentencia & "AND DATE_FORMAT(a.fehtcb, '%Y') <='" & CLng(Right(cboEjercicio.Text, 4)) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.fehtcb=a.fehtcb) "
        sSentencia = sSentencia & "ORDER BY a.fehtcb"
       Case 7     ' estado financieros
        ' Proyectos
        sSentencia = "INSERT INTO codpe (coddpe, detdpe, detdpex, codcco, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT a.coddpe, a.detdpe, a.detdpex, a.codcco, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM codpe a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM codpe b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.coddpe=a.coddpe) "
        sSentencia = sSentencia & "ORDER BY a.coddpe"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' cabecera de estados financieros
        sSentencia = "INSERT INTO " & sTabla & " (" & Replace(sColumnas, "a.", "") & sWhere & ", codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codefi=a.codefi) "
        sSentencia = sSentencia & "ORDER BY a.codefi"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' Detalle de estado financieros
        sSentencia = "INSERT INTO coefilin (codefi, nrolin, detlin, detlinx, tpolin, fmllin, bsepct, grppct, imp1, pct1, imp2, pct2, indlat, indbdesup, indbdeinf, indfondet, indfondet_syd, indfonimp, pdoano, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT a.codefi, a.nrolin, a.detlin, a.detlinx, a.tpolin, a.fmllin, a.bsepct, a.grppct, a.imp1, a.pct1, a.imp2, a.pct2, a.indlat, a.indbdesup, a.indbdeinf, a.indfondet, a.indfondet_syd, a.indfonimp, '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM coefilin a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM coefilin b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codefi=a.codefi) "
        sSentencia = sSentencia & "ORDER BY a.codefi"
       Case 8     ' asiento tipo
        ' cabecera
        sSentencia = "INSERT INTO " & sTabla & " (" & Replace(sColumnas, "a.", "") & sWhere & ", codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codasi=a.codasi) "
        sSentencia = sSentencia & "ORDER BY a.codasi"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' detalle
        sSentencia = "INSERT INTO coasidet (codasi, tpocnc, codcta_mn, orden, codcta_me, codcco, pordst, pdoano, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT a.codasi, a.tpocnc, a.codcta_mn, a.orden, a.codcta_me, a.codcco, a.pordst, '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM coasidet a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM coasidet b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codasi=a.codasi) "
        sSentencia = sSentencia & "ORDER BY a.codasi"
       Case 9     ' flujo de caja y otras tablas
        ' flujo de efectivo
        sSentencia = sSentencia & "SELECT " & sColumnas & ", '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM " & sTabla & " a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codefe=a.codefe) "
        sSentencia = sSentencia & "ORDER BY a.codefe"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' flujo de caja
        sSentencia = "INSERT INTO cofjo (codfjo, detfjo, detfjox, tpofjo, codefe, pdoano, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT a.codfjo, a.detfjo, a.detfjox, a.tpofjo, a.codefe, '" & gsAnoAct & "' AS pdoano, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM cofjo a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM cofjo b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "' "
        sSentencia = sSentencia & "AND b.codfjo=a.codfjo) "
        sSentencia = sSentencia & "ORDER BY a.codfjo"
        pocnnMain.Execute sSentencia, nNumRegistros
        ' medio de pago
        sSentencia = "INSERT INTO bnmediopago (codmed, desmed, abvmed, indmod, estadomed, codemp, usrcre, fyhcre) "
        sSentencia = sSentencia & "SELECT a.codmed, a.desmed, a.abvmed, a.indmod, a.estadomed, '" & gsCodEmp & "' AS codemp, "
        sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM bnmediopago a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM bnmediopago b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.codmed=a.codmed) "
        sSentencia = sSentencia & "ORDER BY a.codmed"
        
        
'ini 2015-09-17 tabla configu. parametros
        pocnnMain.Execute sSentencia, nNumRegistros
        ' cofigura cocfg
        sSentencia = "DELETE FROM cocfg  "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        pocnnMain.Execute sSentencia, nNumRegistros
        
        sSentencia = "INSERT INTO cocfg (" & flst_field("cocfg") & ") "
        sSentencia = sSentencia & "SELECT " & flst_field_sele("cocfg") & " "
        'sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM cocfg a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM cocfg b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "') "
        'sSentencia = sSentencia & "AND b.codfjo=a.codfjo) "
        ''sSentencia = sSentencia & "ORDER BY a.codfjo"
        pocnnMain.Execute sSentencia, nNumRegistros
        
        ' cofigura cocfg
        sSentencia = "DELETE FROM tgcfg  "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
        pocnnMain.Execute sSentencia, nNumRegistros
        
        sSentencia = "INSERT INTO tgcfg (" & flst_field("tgcfg") & ") "
        sSentencia = sSentencia & "SELECT " & flst_field_sele("tgcfg") & " "
        'sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM tgcfg a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND a.pdoano='" & Right(cboEjercicio.Text, 4) & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM tgcfg b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.pdoano='" & gsAnoAct & "') "
        'sSentencia = sSentencia & "AND b.codfjo=a.codfjo) "
        ''sSentencia = sSentencia & "ORDER BY a.codfjo"
        pocnnMain.Execute sSentencia, nNumRegistros
        
        'detraccion
        sSentencia = "INSERT INTO codetrac (" & flst_field("codetrac") & ") "
        sSentencia = sSentencia & "SELECT " & flst_field_sele("codetrac") & " "
        'sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
        sSentencia = sSentencia & "FROM codetrac a "
        sSentencia = sSentencia & "WHERE a.codemp='" & txtDato(0).Text & "' "
        sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM codetrac b "
        sSentencia = sSentencia & "WHERE b.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND b.coddetrac=a.coddetrac) "
        sSentencia = sSentencia & "ORDER BY a.coddetrac"
       
'fin 2015-09-17 tabla configu. parametros
        
      End Select
      pocnnMain.Execute sSentencia, nNumRegistros
    End If
    pgbProgreso(0).Value = nContador + 1
  Next nContador

End Sub
Function flst_field(xaTabla As String) As String
    Dim yyfield As String
    Dim yrst As ADODB.Recordset
    Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, xaTabla))
    With yrst
        If .RecordCount > 0 Then .MoveFirst
        yyfield = yyfield & !COLUMN_NAME '+ ","
        .MoveNext
        Do While Not .EOF
            yyfield = yyfield & "," & !COLUMN_NAME
            .MoveNext
        Loop
    End With
    fRstClose yrst
    flst_field = yyfield
End Function
Function flst_field_sele(xaTabla As String) As String
    Dim yyfield As String
    Dim yrst As ADODB.Recordset
    Set yrst = fRstOpenBuscar(pocnnMain, yrst, finformation_schema_COLUMNS2(gsNomBDS, xaTabla))
    With yrst
        If .RecordCount > 0 Then .MoveFirst
        Select Case UCase(!COLUMN_NAME)
        Case UCase("codemp")
            yyfield = "'" & gsCodEmp & "' AS codemp"
        Case UCase("pdoano")
            yyfield = "'" & gsAnoAct & "' AS pdoano"
        Case Else
            yyfield = "a." & !COLUMN_NAME  '+ ","
        End Select
        .MoveNext
        Do While Not .EOF
            Select Case UCase(!COLUMN_NAME)
            Case UCase("codemp")
                yyfield = yyfield & ",'" & gsCodEmp & "' AS codemp"
            Case UCase("pdoano")
                yyfield = yyfield & ",'" & gsAnoAct & "' AS pdoano"
            Case UCase("usrcre")
                yyfield = yyfield & ",'" & gsAbvUsr & "' AS usrcre"
            Case UCase("fyhcre")
                yyfield = yyfield & ",'" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre"
            Case UCase("fyhmdf"), UCase("usrmdf")
                yyfield = yyfield & ",null AS fyhmdf"
            Case Else
                yyfield = yyfield & ",a." & !COLUMN_NAME
            End Select
            
            .MoveNext
        Loop
    End With
    fRstClose yrst
    flst_field_sele = yyfield
End Function

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0                           'Cambiar (añadir índices).
    modAyuBus.Emp_Usu "oxu.codusr='" & gsCodUsr & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + frmTablas(3).Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + frmTablas(3).Left + txtDato(tnIndex).Left
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
    With porstEmpresa
      .MoveFirst
      .Find "codemp='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!RazEmp), "", !RazEmp)
      End If
    End With
  End Select
End Function

Private Sub ppCentraliza_Proceso()
  Static sSentencia As String, sMilinea As String
  Static nContador As Integer, nArchivo As Integer
  Static nRegistro As Long, nNumRegistros As Long
  Static sArchivo As String, sCaracter As String
  Static sTabla As String, sColumnas As String
  Static sWhere As String, s_Sentencia As String
  Static n_Posicion1 As Integer, n_Posicion2 As Integer
  Static sFecha As String, sFecFinMes As String, sFecIniMes As String
  Static porstTmp As ADODB.Recordset
  Dim nOrden As Long, sCampoDro As String
  Dim sDiario As String, sComprobante As String
  
  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  sCaracter = "|"
  sFecIniMes = Format("01/" & gsMesAct & "/" & gsAnoAct, "dd/mm/yyyy")
  sFecFinMes = gfUltDia(Format("01/" & gsMesAct & "/" & gsAnoAct, "dd/mm/yyyy"))
  sCampoDro = IIf(Val(cmbParametro(1).Text) > 1, "coddro2", "coddro1")
  
  pgbProgreso(0).Max = chkCentraProceso.Count
  pgbProgreso(0).Value = pgbProgreso(0).Min
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkCentraProceso.Count - 1
    ' Verifico que se haya seleccionado
    If (chkCentraProceso(nContador).Value = vbChecked) Then
      If nContador = 0 Or nContador = 2 Then      ' Registro de compras, honarios
        ' Genero la tabla temporal e inserto los registros
        sSentencia = " SELECT DISTINCT cab.cprcnroanno AS anopvs, cab.cprccodtipd AS tipint, det.cpddnrodoci AS nroint,"
        sSentencia = sSentencia & " det.cpddcodprv AS codaux, det.cpddcodtipd AS tipdoc, det.cpddnrodoce AS nrodoc,"
        sSentencia = sSentencia & " det.cpddfecrec AS fecope, det.cpddfecdoc AS fecdoc, det.cpddfecvto AS fecven,"
        sSentencia = sSentencia & " det.cpddfecrec AS fecref, (CASE det.cpddcodmon WHEN 'S/.' THEN 'N' ELSE 'E' END) AS tipmon,"
        sSentencia = sSentencia & " 1.1111 AS tipcam, det.cpddnrodocr AS docref, det.cpdddesobs AS glosa, cab.cprcnroper AS mespvs,"
        sSentencia = sSentencia & " cab.cprctipcmp AS tpocpb, cab.cprcnrocmp AS nrocpb, ' ' AS nrodetrac, det.cpddfecrec AS fecdetrac,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'S/.' THEN det.cpddimpbasi ELSE cab.cprcimpbasi END, 0.00)) AS impbasgrmn,"
        sSentencia = sSentencia & " 0.00 AS impbasgnmn, 0.00 AS impbasngmn,"
        sSentencia = sSentencia & " NVL(ROUND(CASE det.cpddcodmon WHEN 'S/.' THEN (ABS(NVL(det.cpddimpnet, 0)) - ABS(NVL(det.cpddimpbasi, 0))) ELSE (ABS(NVL(cab.cprcimpnet, 0)) - ABS(NVL(cab.cprcimpbasi, 0))) END, 2), 0.00) AS impexomn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'S/.' THEN det.cpddimpigv ELSE cab.cprcimpigv END, 0.00)) AS impigvtomn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'S/.' THEN det.cpddimpigv ELSE cab.cprcimpigv END, 0.00)) AS impigvgrmn,"
        sSentencia = sSentencia & " 0.00 AS impigvgnmn, 0.00 AS impigvngmn, 0.00 AS impiscmn, 0.00 AS impotrmn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'S/.' THEN det.cpddimptot ELSE cab.cprcimptot END, 0.00)) AS imptotmn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'US$' THEN det.cpddimpbasi ELSE 0.00 END, 0.00)) AS impbasgrme,"
        sSentencia = sSentencia & " 0.00 AS impbasgnme, 0.00 AS impbasngme,"
        sSentencia = sSentencia & " NVL(ROUND(CASE det.cpddcodmon WHEN 'US$' THEN (ABS(NVL(det.cpddimpnet, 0)) - ABS(NVL(det.cpddimpbasi, 0))) ELSE 0.00 END, 2), 0.00) AS impexome,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'US$' THEN det.cpddimpigv ELSE 0.00 END, 0.00)) AS impigvtome,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'US$' THEN det.cpddimpigv ELSE 0.00 END, 0.00)) AS impigvgrme,"
        sSentencia = sSentencia & " 0.00 AS impigvgnme, 0.00 AS impigvngme, 0.00 AS impiscme, 0.00 AS impotrme,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.cpddcodmon WHEN 'US$' THEN det.cpddimptot ELSE 0.00 END, 0.00)) AS imptotme,"
        sSentencia = sSentencia & " cab.cprccodest AS anucab, det.cpddcodest AS anudet,"
        sSentencia = sSentencia & " 'td' AS tipodocu, 'ldro' AS diario, 'cenco' AS cencosto"
        sSentencia = sSentencia & " FROM cpddrdc cab, cpdddocu det"
        sSentencia = sSentencia & " WHERE cab.cprccodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND cab.cprccodsuc='" & cmbParametro(1).Text & "'"
        sSentencia = sSentencia & " AND cab.cprcnroanno=" & gsAnoAct
        sSentencia = sSentencia & " AND cab.cprcnroper=" & gsMesAct
        sSentencia = sSentencia & " AND det.cpddcodcia=cab.cprccodcia"
        sSentencia = sSentencia & " AND det.cpddcodsuc=cab.cprccodsuc"
        sSentencia = sSentencia & " AND det.cpddcodtipd=cab.cprccodtipd"
        sSentencia = sSentencia & " AND det.cpddnrodoci=cab.cprcnrodoci"
        sSentencia = sSentencia & " ORDER BY tipint, nroint"
        
        ' Elimino si existe tabla
        s_Sentencia = "SELECT COUNT(*) AS nExiste FROM dba_tables WHERE table_name='" & UCase(IIf(nContador = 0, "tmp_regcom", "tmp_reghpr")) & "'"
        With porstTmp
          If .State = adStateOpen Then .Close
          .ActiveConnection = pocnnMain
          .Source = s_Sentencia
          .CursorType = adOpenDynamic
          .LockType = adLockReadOnly
          .Open
          nRegistro = !nExiste
          .Close
        End With
        If nRegistro = 1 Then
          s_Sentencia = "DROP TABLE " & IIf(nContador = 0, "tmp_regcom", "tmp_reghpr")
          pocnnMain.Execute s_Sentencia, nRegistro
        End If
        
        s_Sentencia = "CREATE GLOBAL TEMPORARY TABLE " & IIf(nContador = 0, "tmp_regcom", "tmp_reghpr") & " AS"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        If nContador = 0 Then   ' Registro de compras
          s_Sentencia = "INSERT INTO tmp_regcom"
          s_Sentencia = s_Sentencia & sSentencia
          pocnnMain.Execute s_Sentencia
          
          ' Actualizo el tipo de cambio
          s_Sentencia = "UPDATE tmp_regcom"
          s_Sentencia = s_Sentencia & " SET tipcam=(SELECT tpc.cntpvalvta"
          s_Sentencia = s_Sentencia & " FROM cntipcam tpc"
          s_Sentencia = s_Sentencia & " WHERE tpc.cntpfecdia=tmp_regcom.fecdoc"
          s_Sentencia = s_Sentencia & " AND tpc.cntpcodmon='" & gsTpoMon_Sgn_ME & "')"
          pocnnMain.Execute s_Sentencia
          ' Actualizo los importes en dolares
          s_Sentencia = "UPDATE tmp_regcom"
          s_Sentencia = s_Sentencia & " SET impbasgrme=ROUND(impbasgrmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impbasgnme=ROUND(impbasgnmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impbasngme=ROUND(impbasngmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impexome=ROUND(impexomn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impigvtome=ROUND(impigvtomn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impigvgrme=ROUND(impigvgrmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impigvgnme=ROUND(impigvgnmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impigvngme=ROUND(impigvngmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impiscme=ROUND(impiscmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " impotrme=ROUND(impotrmn/tipcam, 2),"
          s_Sentencia = s_Sentencia & " imptotme=ROUND(imptotmn/tipcam, 2)"
          s_Sentencia = s_Sentencia & " WHERE tipmon='" & TPOMON_NAC & "'"
          pocnnMain.Execute s_Sentencia
        End If
        
        ' [Actualizar tipo de documento, diario, centro de costo
        s_Sentencia = "UPDATE " & IIf(nContador = 0, "tmp_regcom", "tmp_reghpr")
        s_Sentencia = s_Sentencia & " SET tipodocu=(SELECT tip.codtpodoc FROM tmptpodoc tip"
        s_Sentencia = s_Sentencia & " WHERE tip.cntbkeyocur=tmp_regcom.tipdoc)"
        pocnnMain.Execute s_Sentencia
        
        s_Sentencia = "UPDATE " & IIf(nContador = 0, "tmp_regcom", "tmp_reghpr")
        s_Sentencia = s_Sentencia & " SET diario=(SELECT dro." & sCampoDro & " FROM tmpdro dro"
        s_Sentencia = s_Sentencia & " WHERE dro.cntbkeyocur=tmp_regcom.tpocpb)"
        pocnnMain.Execute s_Sentencia
        ']
        
        ' Campos de seleccion
        s_Sentencia = "tipint, nroint, codaux, tipdoc, tipodocu, nrodoc, fecope, fecdoc, fecven, fecref, tipmon,"
        s_Sentencia = s_Sentencia & " tipcam, docref, glosa, mespvs, tpocpb, diario, nrocpb, nrodetrac, fecdetrac,"
        s_Sentencia = s_Sentencia & " impbasgrmn, impbasgnmn,impbasngmn, impexomn, impigvtomn,"
        s_Sentencia = s_Sentencia & " impigvgrmn, impigvgnmn, impigvngmn, impiscmn,impotrmn, imptotmn,"
        s_Sentencia = s_Sentencia & " impbasgrme, impbasgnme, impbasngme, impexome, impigvtome,"
        s_Sentencia = s_Sentencia & " impigvgrme, impigvgnme, impigvngme, impiscme, impotrme, imptotme,"
        s_Sentencia = s_Sentencia & " anucab, anudet"
      ElseIf nContador = 1 Then   ' Registro de Ventas
        ' Genero la tabla temporal e inserto los registros
        sSentencia = " SELECT DISTINCT cab.cprcnroanno AS anopvs, cab.cprccodtipd AS tipint, det.clddnrodoci AS nroint,"
        sSentencia = sSentencia & " det.clddcodtipd AS tipdoc, det.clddnrodoce AS nrodoc, det.clddfecent AS fecope,"
        sSentencia = sSentencia & " ' ' AS bolfin, det.clddcodcli AS codaux, det.clddfecdoc AS fecdoc,"
        sSentencia = sSentencia & " det.clddfecvto AS fecven, det.clddfecing AS fecref,"
        sSentencia = sSentencia & " (CASE det.clddcodmon WHEN 'S/.' THEN 'N' ELSE 'E' END) AS tipmon,"
        sSentencia = sSentencia & " 1.1111 AS tipcam, det.clddnrodocr AS docref, det.cldddesobs AS glosa,"
        sSentencia = sSentencia & " cab.cprcnroper AS mespvs, cab.cprctipcmp AS tpocpb, cab.cprcnrocmp AS nrocpb,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.clddcodmon WHEN 'S/.' THEN det.clddimpbasi ELSE cab.cprcimpbasi END, 0.00)) AS impbasgrmn,"
        sSentencia = sSentencia & " 0.00 AS impbasngmn,"
        sSentencia = sSentencia & " NVL(ROUND(CASE det.clddcodmon WHEN 'S/.' THEN (ABS(NVL(det.clddimpnet, 0)) - ABS(NVL(det.clddimpbasi, 0))) ELSE (ABS(NVL(cab.cprcimpnet, 0)) - ABS(NVL(cab.cprcimpbasi, 0))) END, 2), 0.00) AS impexomn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.clddcodmon WHEN 'S/.' THEN det.clddimpigv ELSE cab.cprcimpigv END, 0.00)) AS impigvgrmn,"
        sSentencia = sSentencia & " 0.00 AS impiscmn, 0.00 AS impotrmn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.clddcodmon WHEN 'S/.' THEN det.clddimptot ELSE cab.cprcimptot END, 0.00)) AS imptotmn,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.clddcodmon WHEN 'US$' THEN det.clddimpbasi ELSE 0.00 END, 0.00)) AS impbasgrme,"
        sSentencia = sSentencia & " 0.00 AS impbasngme,"
        sSentencia = sSentencia & " NVL(ROUND(CASE det.clddcodmon WHEN 'US$' THEN (ABS(NVL(det.clddimpnet, 0)) - ABS(NVL(det.clddimpbasi, 0))) ELSE 0.00 END, 2), 0.00) AS impexome,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.clddcodmon WHEN 'US$' THEN det.clddimpigv ELSE 0.00 END, 0.00)) AS impigvgrme,"
        sSentencia = sSentencia & " 0.00 AS impiscme, 0.00 AS impotrme,"
        sSentencia = sSentencia & " ABS(NVL(CASE det.clddcodmon WHEN 'US$' THEN det.clddimptot ELSE 0.00 END, 0.00)) AS imptotme,"
        sSentencia = sSentencia & " cab.cprccodest AS anucab, det.clddcodest AS anudet,"
        sSentencia = sSentencia & " 'td' AS tipodocu, 'ldro' AS diario, 'cenco' AS cencosto"
        sSentencia = sSentencia & " FROM clddrdv cab, cldddocu det"
        sSentencia = sSentencia & " WHERE cab.cprccodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND cab.cprccodsuc='" & cmbParametro(1).Text & "'"
        sSentencia = sSentencia & " AND cab.cprcnroanno=" & gsAnoAct
        sSentencia = sSentencia & " AND cab.cprcnroper=" & gsMesAct
        sSentencia = sSentencia & " AND det.clddcodcia=cab.cprccodcia"
        sSentencia = sSentencia & " AND det.clddcodsuc=cab.cprccodsuc"
        sSentencia = sSentencia & " AND det.clddcodtipd=cab.cprccodtipd"
        sSentencia = sSentencia & " AND det.clddnrodoci=cab.cprcnrodoci"
        sSentencia = sSentencia & " ORDER BY tipint, nroint"
        
        ' Elimino si existe tabla
        s_Sentencia = "SELECT COUNT(*) AS nExiste FROM dba_tables WHERE table_name='TMP_REGVTA'"
        With porstTmp
          If .State = adStateOpen Then .Close
          .ActiveConnection = pocnnMain
          .Source = s_Sentencia
          .CursorType = adOpenDynamic
          .LockType = adLockReadOnly
          .Open
          nRegistro = !nExiste
          .Close
        End With
        If nRegistro = 1 Then
          s_Sentencia = "DROP TABLE tmp_regvta"
          pocnnMain.Execute s_Sentencia, nRegistro
        End If
        
        s_Sentencia = "CREATE GLOBAL TEMPORARY TABLE tmp_regvta AS"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "INSERT INTO tmp_regvta"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        
        ' Actualizo el tipo de cambio
        s_Sentencia = "UPDATE tmp_regvta"
        s_Sentencia = s_Sentencia & " SET tipcam=(SELECT tpc.cntpvalvta"
        s_Sentencia = s_Sentencia & " FROM cntipcam tpc"
        s_Sentencia = s_Sentencia & " WHERE tpc.cntpfecdia=tmp_regvta.fecdoc"
        s_Sentencia = s_Sentencia & " AND tpc.cntpcodmon='" & gsTpoMon_Sgn_ME & "')"
        pocnnMain.Execute s_Sentencia
        ' Actualizo los importes en dolares
        s_Sentencia = "UPDATE tmp_regvta"
        s_Sentencia = s_Sentencia & " SET impbasgrme=ROUND(impbasgrmn/tipcam, 2),"
        s_Sentencia = s_Sentencia & " impbasngme=ROUND(impbasngmn/tipcam, 2),"
        s_Sentencia = s_Sentencia & " impexome=ROUND(impexomn/tipcam, 2),"
        s_Sentencia = s_Sentencia & " impigvgrme=ROUND(impigvgrmn/tipcam, 2),"
        s_Sentencia = s_Sentencia & " impiscme=ROUND(impiscmn/tipcam, 2),"
        s_Sentencia = s_Sentencia & " impotrme=ROUND(impotrmn/tipcam, 2),"
        s_Sentencia = s_Sentencia & " imptotme=ROUND(imptotmn/tipcam, 2)"
        s_Sentencia = s_Sentencia & " WHERE tipmon='" & TPOMON_NAC & "'"
        pocnnMain.Execute s_Sentencia
        
        ' [Actualizar tipo de documento, diario, centro de costo
        s_Sentencia = "UPDATE tmp_regvta"
        s_Sentencia = s_Sentencia & " SET tipodocu=(SELECT tip.codtpodoc FROM tmptpodoc tip"
        s_Sentencia = s_Sentencia & " WHERE tip.cntbkeyocur=tmp_regvta.tipdoc)"
        pocnnMain.Execute s_Sentencia
        
        s_Sentencia = "UPDATE tmp_regvta"
        s_Sentencia = s_Sentencia & " SET diario=(SELECT dro." & sCampoDro & " FROM tmpdro dro"
        s_Sentencia = s_Sentencia & " WHERE dro.cntbkeyocur=tmp_regvta.tpocpb)"
        pocnnMain.Execute s_Sentencia
        ']
        
        ' Campos de seleccion
        s_Sentencia = "tipint, nroint, tipdoc, tipodocu, nrodoc, fecope, bolfin, codaux, fecdoc, fecven, fecref, tipmon,"
        s_Sentencia = s_Sentencia & " tipcam, docref, glosa, mespvs, tpocpb, diario, nrocpb,"
        s_Sentencia = s_Sentencia & " impbasgrmn, impbasngmn, impexomn, impigvgrmn,"
        s_Sentencia = s_Sentencia & " impiscmn,impotrmn, imptotmn,"
        s_Sentencia = s_Sentencia & " impbasgrme, impbasngme, impexome, impigvgrme,"
        s_Sentencia = s_Sentencia & " impiscme, impotrme, imptotme,"
        s_Sentencia = s_Sentencia & " anucab, anudet"
      ElseIf nContador = 3 Then   ' registro de diario
        ' Genero la tabla temporal e inserto los registros
        sSentencia = " SELECT det.ctdccodtipc AS coddro, det.ctdcnrocomp AS nrocpb, det.ctdcnroper mespvs,"
        sSentencia = sSentencia & " cab.ctccfecing AS feccon, cab.ctccdesglo AS glocpb, 'D' AS tpocpb,"
        sSentencia = sSentencia & " det.ctdcnrolin AS orden, det.ctdcnrolin AS ordendst, det.ctdccodtipd AS tpodoc,"
        sSentencia = sSentencia & " det.ctdccodcta AS codcta, det.ctdccodcco AS codcco, '           ' AS codaux,"
        sSentencia = sSentencia & " '          ' AS nrodocu, cab.ctccfecing AS fecdoc, cab.ctccfecing AS fecven,"
        sSentencia = sSentencia & " cab.ctccfecing AS fecrec, det.ctdcnrodoci AS docref, cab.ctccdesglo AS glosadet,"
        sSentencia = sSentencia & " (CASE ctdcflgdeb  WHEN 'D' THEN 'D' ELSE 'H' END) AS DebHab,"
        sSentencia = sSentencia & " 'O' AS tpodet, (CASE det.ctdccodmono WHEN 'S/.' THEN 'N' ELSE 'E' END) AS tpomon,"
        sSentencia = sSentencia & " 'V' AS tipcam, ABS(NVL((CASE det.ctdccodmono WHEN 'S/.' THEN det.ctdcimptasd ELSE det.ctdcimptass END), 1.0000)) AS imptipcam,"
        sSentencia = sSentencia & " ABS(NVL(ctdcims, 0.00)) AS impmn, ABS(NVL(ctdcimd, 0.00)) AS impme,"
        sSentencia = sSentencia & " cab.ctcccodest AS anucab, det.ctdccodest AS anudet,"
        sSentencia = sSentencia & " 'td' AS tipodocu, 'ldro' AS diario, 'cenco' AS cencosto"
        sSentencia = sSentencia & " FROM ctdicomb det, ctdccomb cab"
        sSentencia = sSentencia & " WHERE det.ctdccodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND det.ctdccodsuc='" & cmbParametro(1).Text & "'"
        sSentencia = sSentencia & " AND det.ctdcnroanno=" & gsAnoAct
        sSentencia = sSentencia & " AND det.ctdcnroper=" & gsMesAct
        sSentencia = sSentencia & " AND cab.ctcccodcia=det.ctdccodcia"
        sSentencia = sSentencia & " AND cab.ctcccodsuc=det.ctdccodsuc"
        sSentencia = sSentencia & " AND cab.ctccnroanno=det.ctdcnroanno"
        sSentencia = sSentencia & " AND cab.ctccnroper=det.ctdcnroper"
        sSentencia = sSentencia & " AND cab.ctcccodtipc=det.ctdccodtipc"
        sSentencia = sSentencia & " AND cab.ctccnrocomp=det.ctdcnrocomp"
        sSentencia = sSentencia & " AND cab.ctcccodest IN('0', '1', '2')"
        sSentencia = sSentencia & " ORDER BY coddro, nrocpb, orden"
        
        ' Elimino si existe tabla
        s_Sentencia = "SELECT COUNT(*) AS nExiste FROM dba_tables WHERE table_name='TMP_REGDRO'"
        With porstTmp
          If .State = adStateOpen Then .Close
          .ActiveConnection = pocnnMain
          .Source = s_Sentencia
          .CursorType = adOpenDynamic
          .LockType = adLockReadOnly
          .Open
          nRegistro = !nExiste
          .Close
        End With
        If nRegistro = 1 Then
          s_Sentencia = "DROP TABLE tmp_regdro"
          pocnnMain.Execute s_Sentencia, nRegistro
        End If
        
        s_Sentencia = "CREATE TABLE tmp_regdro AS"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        
        ' Actualizo los tipos de comprobantes V:enta y tipo de detalle P:rovision,
        s_Sentencia = "UPDATE tmp_regdro SET tpocpb='V', tpodet='P'"
        s_Sentencia = s_Sentencia & " WHERE EXISTS(SELECT * FROM clddrdv doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cprccodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cprccodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cprcnroanno=" & gsAnoAct
        s_Sentencia = s_Sentencia & " AND doc.cprcnroper=tmp_regdro.mespvs"
        s_Sentencia = s_Sentencia & " AND doc.cprctipcmp=tmp_regdro.coddro"
        s_Sentencia = s_Sentencia & " AND doc.cprcnrocmp=tmp_regdro.nrocpb)"
        pocnnMain.Execute s_Sentencia
        ' Actualizo los tipos de comprobantes C:ompra y tipo de detalle P:rovision,
        s_Sentencia = "UPDATE tmp_regdro SET tpocpb='C', tpodet='P'"
        s_Sentencia = s_Sentencia & " WHERE EXISTS(SELECT * FROM cpddrdc doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cprccodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cprccodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cprcnroanno=" & gsAnoAct
        s_Sentencia = s_Sentencia & " AND doc.cprcnroper=tmp_regdro.mespvs"
        s_Sentencia = s_Sentencia & " AND doc.cprctipcmp=tmp_regdro.coddro"
        s_Sentencia = s_Sentencia & " AND doc.cprcnrocmp=tmp_regdro.nrocpb)"
        pocnnMain.Execute s_Sentencia
        
        ' Actualizo el auxiliar y documento de clientes
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET codaux=("
        s_Sentencia = s_Sentencia & " SELECT doc.clddcodcli FROM cldddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.clddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.clddnrodoci=tmp_regdro.docref)"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET nrodocu=(SELECT doc.clddnrodoce FROM cldddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.clddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.clddnrodoci=tmp_regdro.docref)"
        pocnnMain.Execute s_Sentencia
        ' Actualizo el auxiliar y documento de proveedores
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET codaux=NVL(("
        s_Sentencia = s_Sentencia & " SELECT doc.cpddcodprv FROM cpdddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cpddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.cpddnrodoci=tmp_regdro.docref), tmp_regdro.codaux)"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET nrodocu=NVL((SELECT doc.cpddnrodoce FROM cpdddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cpddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.cpddnrodoci=tmp_regdro.docref), tmp_regdro.nrodocu)"
        pocnnMain.Execute s_Sentencia

        ' Actualizo las fechas de documentos de clientes
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET fecdoc=NVL((SELECT doc.clddfecdoc FROM cldddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.clddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND tmp_regdro.tpodet='" & TPOPVS_PVS & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.clddnrodoci=tmp_regdro.docref), tmp_regdro.fecdoc)"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET fecven=NVL((SELECT doc.clddfecvto FROM cldddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.clddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND tmp_regdro.tpodet='" & TPOPVS_PVS & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.clddnrodoci=tmp_regdro.docref), tmp_regdro.fecven)"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET fecrec=NVL((SELECT doc.clddfecent FROM cldddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.clddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND tmp_regdro.tpodet='" & TPOPVS_PVS & "'"
        s_Sentencia = s_Sentencia & " AND doc.clddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.clddnrodoci=tmp_regdro.docref), tmp_regdro.fecrec)"
        pocnnMain.Execute s_Sentencia
        ' Actualizo las fechas de documentos de proveedores
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET fecdoc=NVL((SELECT doc.cpddfecdoc FROM cpdddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cpddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND tmp_regdro.tpodet='" & TPOPVS_PVS & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.cpddnrodoci=tmp_regdro.docref), tmp_regdro.fecdoc)"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET fecven=NVL((SELECT doc.cpddfecvto FROM cpdddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cpddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND tmp_regdro.tpodet='" & TPOPVS_PVS & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.cpddnrodoci=tmp_regdro.docref), tmp_regdro.fecven)"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET fecrec=NVL((SELECT doc.cpddfecrec FROM cpdddocu doc"
        s_Sentencia = s_Sentencia & " WHERE doc.cpddcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND tmp_regdro.tpodet='" & TPOPVS_PVS & "'"
        s_Sentencia = s_Sentencia & " AND doc.cpddcodtipd=tmp_regdro.tpodoc"
        s_Sentencia = s_Sentencia & " AND doc.cpddnrodoci=tmp_regdro.docref), tmp_regdro.fecrec)"
        pocnnMain.Execute s_Sentencia
        
        ' [Actualizar tipo de documento, diario, centro de costo
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET tipodocu=(SELECT tip.codtpodoc FROM tmptpodoc tip"
        s_Sentencia = s_Sentencia & " WHERE tip.cntbkeyocur=tmp_regdro.tpodoc)"
        pocnnMain.Execute s_Sentencia
        
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET diario=(SELECT dro." & sCampoDro & " FROM tmpdro dro"
        s_Sentencia = s_Sentencia & " WHERE dro.cntbkeyocur=tmp_regdro.coddro"
        s_Sentencia = s_Sentencia & " AND dro.trfdro='1')"
        pocnnMain.Execute s_Sentencia
        
        s_Sentencia = "UPDATE tmp_regdro"
        s_Sentencia = s_Sentencia & " SET cencosto=(SELECT cco.codcencos FROM tmpcencos cco"
        s_Sentencia = s_Sentencia & " WHERE cco.ccm_ccosto=tmp_regdro.codcco)"
        pocnnMain.Execute s_Sentencia
        
        ' Elimino los registro sin diario
        s_Sentencia = "DELETE FROM tmp_regdro"
        s_Sentencia = s_Sentencia & " WHERE diario IS NULL"
        pocnnMain.Execute s_Sentencia
        ' ]
        
        ' Campos de seleccion
        s_Sentencia = "coddro, nrocpb, diario, mespvs, feccon, glocpb, tpocpb, orden, ordendst,"
        s_Sentencia = s_Sentencia & " tpodoc, tipodocu, codcta, codcco, cencosto, codaux, nrodocu, fecdoc,"
        s_Sentencia = s_Sentencia & " fecven, fecrec, docref, glosadet, debhab, tpodet,"
        s_Sentencia = s_Sentencia & " tpomon, tipcam, imptipcam , impmn, impme,"
        s_Sentencia = s_Sentencia & " anucab, anudet"
      End If
      ' Obtengo el archivo de texto
      sArchivo = Choose(nContador + 1, "rc", "rv", "rh", "rd")
      sColumnas = s_Sentencia
      ' Genero el archivo y abro el recordset temporal
      sTabla = Choose(nContador + 1, "tmp_regcom", "tmp_regvta", "tmp_reghpr", "tmp_regdro")
      sWhere = Choose(nContador + 1, "", "", "", "")
      ' Obtengo el archivo de texto libre
      nArchivo = FreeFile
      sSentencia = " SELECT " & IIf(nContador <> 3, "DISTINCT ", "") & sColumnas
      sSentencia = sSentencia & " FROM " & sTabla
      sSentencia = sSentencia & IIf(sWhere <> "", " WHERE ", "") & sWhere
      sSentencia = sSentencia & " ORDER BY 1, 2"
      With porstTmp
        If .State = adStateOpen Then .Close
        .ActiveConnection = pocnnMain
        .Source = sSentencia
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
      End With
      ' Inicializo la variables adicionales
      nOrden = 0
      sDiario = "": sComprobante = ""
      
      If Not (porstTmp.BOF And porstTmp.EOF) Then
        ' Barro todo el Archivo de Texto y lo grabo en Tabla Temporal
        lblProgreso(1).Caption = Choose(gsIdioma, "Exportando Archivo: ", "Exporting File: ") & Trim(chkCentraProceso(nContador).Caption)
        nNumRegistros = porstTmp.RecordCount
        pgbProgreso(1).Max = nNumRegistros
        pgbProgreso(1).Value = pgbProgreso(1).Min
        nRegistro = 0
        sArchivo = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & sArchivo & gsAnoAct & gsMesAct & ".txt"
        ' Elimino archivo de texto si existe
        If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
        If Dir$(sArchivo, vbNormal) = "" Then
          Open sArchivo For Output Access Write Lock Read Write As #nArchivo
          While Not porstTmp.EOF
            nRegistro = nRegistro + 1
            ' Diseño y grabro la linea en el archivo
            sMilinea = ""
            Select Case nContador
             Case 0, 2
              sMilinea = sMilinea & Trim(porstTmp!codaux) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tipodocu) & sCaracter
              sMilinea = sMilinea & Left(Trim(porstTmp!nrodoc), 3) & sCaracter
              sMilinea = sMilinea & Mid(Trim(porstTmp!nrodoc), 4) & sCaracter
              ' Fecha de operación
              sFecha = IIf((Format(porstTmp!fecope, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd")), sFecFinMes, IIf((Format(porstTmp!fecope, "yyyymmdd") < Format(sFecIniMes, "yyyymmdd")), sFecIniMes, porstTmp!fecope))
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              ' Fecha de documento
              sFecha = IIf(Format(porstTmp!fecdoc, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd"), sFecFinMes, porstTmp!fecdoc)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              ' Fecha de vencimiento
              sFecha = IIf(Format(porstTmp!fecven, "yyyymmdd") < Format(sFecFinMes, "yyyymmdd"), sFecFinMes, porstTmp!fecven)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              ' Fecha de referencia
              sFecha = IIf(Format(porstTmp!fecref, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd"), sFecFinMes, porstTmp!fecref)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              
              sMilinea = sMilinea & porstTmp!tipmon & sCaracter
              sMilinea = sMilinea & Format(porstTmp!tipcam, "###0.0000") & sCaracter
              sMilinea = sMilinea & Format(IIf(gsAnoAct >= 2004, IIf(gsMesAct >= "08", 19, 18), 18), "##0.00") & sCaracter
              sMilinea = sMilinea & Format(2, "##0.00") & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!nroint) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(IIf(IsNull(porstTmp!glosa), "", porstTmp!glosa)) & sCaracter
              sMilinea = sMilinea & Format(Trim(porstTmp!mespvs), "00") & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!diario) & sCaracter
              sMilinea = sMilinea & Format(Trim(porstTmp!NroCpb), "000000") & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!nrodetrac) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!fecdetrac) & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasgrmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasgnmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasngmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impexomn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvtomn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvgrmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvgnmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvngmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impiscmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impotrmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!ImpTotMN), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasgrme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasgnme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasngme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impexome), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvtome), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvgrme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvgnme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvngme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impiscme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impotrme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!ImpTotME), "############0.00") & sCaracter
             Case 1
              sMilinea = sMilinea & Trim(porstTmp!tipodocu) & sCaracter
              sMilinea = sMilinea & Left(Trim(porstTmp!nrodoc), 3) & sCaracter
              sMilinea = sMilinea & Mid(Trim(porstTmp!nrodoc), 4) & sCaracter
              ' Fecha de operación
              sFecha = IIf((Format(porstTmp!fecope, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd")), sFecFinMes, IIf((Format(porstTmp!fecope, "yyyymmdd") < Format(sFecIniMes, "yyyymmdd")), sFecIniMes, porstTmp!fecope))
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              sMilinea = sMilinea & Left(Trim(porstTmp!bolfin), 3) & sCaracter
              sMilinea = sMilinea & Mid(Trim(porstTmp!bolfin), 4) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!codaux) & sCaracter
              ' Fecha de documento
              sFecha = IIf(Format(porstTmp!fecdoc, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd"), sFecFinMes, porstTmp!fecdoc)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              ' Fecha de vencimiento
              sFecha = IIf(Format(porstTmp!fecven, "yyyymmdd") < Format(sFecFinMes, "yyyymmdd"), sFecFinMes, porstTmp!fecven)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tipmon) & sCaracter
              sMilinea = sMilinea & Format(porstTmp!tipcam, "###0.0000") & sCaracter
              sMilinea = sMilinea & Format(IIf(gsAnoAct >= 2004, IIf(gsMesAct >= "08", 19, 18), 18), "##0.00") & sCaracter
              sMilinea = sMilinea & Format(2, "##0.00") & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!nroint) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(IIf(IsNull(porstTmp!glosa), "", porstTmp!glosa)) & sCaracter
              sMilinea = sMilinea & Format(Trim(porstTmp!mespvs), "00") & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!diario) & sCaracter
              sMilinea = sMilinea & Format(Trim(porstTmp!NroCpb), "000000") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasgrmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasngmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impexomn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvgrmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impiscmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impotrmn), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!ImpTotMN), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasgrme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impbasngme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impexome), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impigvgrme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impiscme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!impotrme), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", 0, porstTmp!ImpTotME), "############0.00") & sCaracter
              sMilinea = sMilinea & IIf(porstTmp!anucab = "3" And porstTmp!anudet = "0", "S", "N") & sCaracter
             Case 3
              ' Inicializo las variables
              If Not (Trim(porstTmp!diario) = sDiario And Trim(porstTmp!NroCpb) = sComprobante) Then
                sDiario = Trim(porstTmp!diario)
                sComprobante = Trim(porstTmp!NroCpb)
                nOrden = 0
              End If
              nOrden = nOrden + 1
             ' Configuro la linea de grabación
              sMilinea = sMilinea & Trim(porstTmp!diario) & sCaracter
              sMilinea = sMilinea & Format(Trim(porstTmp!NroCpb), "000000") & sCaracter
              sMilinea = sMilinea & Format(Trim(porstTmp!mespvs), "00") & sCaracter
              ' Fecha de operación
              sFecha = IIf((Format(porstTmp!feccon, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd")), sFecFinMes, IIf((Format(porstTmp!feccon, "yyyymmdd") < Format(sFecIniMes, "yyyymmdd")), sFecIniMes, porstTmp!feccon))
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(IIf(IsNull(porstTmp!glocpb), "", porstTmp!glocpb)) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tpocpb) & sCaracter
              sMilinea = sMilinea & Trim$(nOrden) & sCaracter
              sMilinea = sMilinea & Trim$(nOrden) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tipodocu) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!codcta) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!cencosto) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!codaux) & sCaracter
              sMilinea = sMilinea & Left(Trim(porstTmp!nrodocu), 3) & sCaracter
              sMilinea = sMilinea & Mid(Trim(porstTmp!nrodocu), 4) & sCaracter
              ' Fecha de documento
              sFecha = Format(IIf(IsNull(porstTmp!fecdoc), sFecFinMes, porstTmp!fecdoc), "dd/mm/yyyy")
              sFecha = IIf(Format(sFecha, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd"), sFecFinMes, sFecha)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              ' Fecha de vencimiento
              sFecha = IIf(Format(porstTmp!fecven, "yyyymmdd") < Format(sFecFinMes, "yyyymmdd"), sFecFinMes, porstTmp!fecven)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              ' Fecha de recepcion
              sFecha = Format(IIf(IsNull(porstTmp!fecrec), sFecFinMes, porstTmp!fecrec), "dd/mm/yyyy")
              sFecha = IIf(Format(sFecha, "yyyymmdd") > Format(sFecFinMes, "yyyymmdd"), sFecFinMes, sFecha)
              sMilinea = sMilinea & Format(sFecha, "dd/mm/yyyy") & sCaracter
              
              sMilinea = sMilinea & Trim(porstTmp!tpodoc) & "/" & Trim(porstTmp!docref) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(IIf(IsNull(porstTmp!glosadet), "", porstTmp!glosadet)) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!debhab) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tpodet) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tpomon) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tipcam) & sCaracter
              sMilinea = sMilinea & Format(porstTmp!imptipcam, "###0.0000") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "0", 0, porstTmp!ImpMN), "############0.00") & sCaracter
              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "0", 0, porstTmp!ImpME), "############0.00") & sCaracter
'              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "0" And porstTmp!anudet = "0", 0, porstTmp!ImpMN), "############0.00") & sCaracter
'              sMilinea = sMilinea & Format(IIf(porstTmp!anucab = "0" And porstTmp!anudet = "0", 0, porstTmp!ImpME), "############0.00") & sCaracter
            End Select
            Print #nArchivo, sMilinea
            pgbProgreso(1).Value = nRegistro
            DoEvents
            porstTmp.MoveNext
          Wend
          Close #nArchivo
        End If
      End If
      pocnnMain.Execute "DROP TABLE " & sTabla
      porstTmp.Close
    End If
    pgbProgreso(0).Value = nContador + 1
  Next nContador
  Set porstTmp = Nothing

End Sub

Private Sub ppCentraliza_Tablas()
  Static sSentencia As String, sMilinea As String
  Static nContador As Integer, nArchivo As Integer
  Static nRegistro As Long, nNumRegistros As Long
  Static sArchivo As String, sCaracter As String
  Static sTabla As String, sColumnas As String
  Static sWhere As String, s_Sentencia As String
  Static n_Posicion1 As Integer, n_Posicion2 As Integer
  Static porstTmp As ADODB.Recordset

  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  sCaracter = "|"
  pgbProgreso(0).Max = chkCentraTabla.Count
  pgbProgreso(0).Value = pgbProgreso(0).Min
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkCentraTabla.Count - 1
    ' Verifico que se haya seleccionado
    If (chkCentraTabla(nContador).Value = vbChecked) Then
      If nContador = 1 Then     ' Plan de cuentas
        ' Genero la tabla temporal e inserto los registros
        sSentencia = " SELECT DISTINCT ctctcodcta AS codcta, ctctdescta AS detcta,"
        sSentencia = sSentencia & " (CASE ctctflgacu WHEN 'N' THEN 'D' ELSE 'T' END) AS tpocta,"
        sSentencia = sSentencia & " (CASE ctctcodtip WHEN 'DB' THEN 'D' ELSE 'A' END) AS natcta,"
        sSentencia = sSentencia & " LPAD(ctctcodcta, 1) AS tposdo,"
        sSentencia = sSentencia & " (CASE NVL(ctctcodaux, '') WHEN '' THEN '' ELSE (CASE ctctflgsal WHEN 'S' THEN 'D' ELSE 'A' END) END) AS tpoanl,"
        sSentencia = sSentencia & " ctctcodasrf, ctctcodcta AS codcta_dst_deb, ctctcodcta AS codcta_dst_hab,"
        sSentencia = sSentencia & " (CASE ctctflgmon WHEN 2 THEN 'E' ELSE 'N' END) AS tpomon, 'V' AS tpotcb, 'N' AS tpoajd,"
        sSentencia = sSentencia & " ctctcodcta AS codcta_ajd_deb, ctctcodcta AS codcta_ajd_hab,"
        sSentencia = sSentencia & " (CASE ctctactsal WHEN '1' THEN 'S' ELSE 'N' END)AS indajd,"
        sSentencia = sSentencia & " (CASE ctctflgcc WHEN 'S' THEN 'S' ELSE 'N' END) AS indcco,"
        sSentencia = sSentencia & " (CASE ctctflgsal WHEN 'S' THEN 'S' ELSE 'N' END) AS inddoc,"
        sSentencia = sSentencia & " 'N' AS indmoe, (CASE NVL(ctctflgpsto, '') WHEN '' THEN 'N' ELSE 'S' END) AS indpsp"
        sSentencia = sSentencia & " FROM ctdmctas"
        sSentencia = sSentencia & " WHERE ctctcodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND ctctcodsuc='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " ORDER BY codcta"
        s_Sentencia = "CREATE GLOBAL TEMPORARY TABLE tbltmp_cuentas AS"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "INSERT INTO tbltmp_cuentas"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        
        ' Actualizo la cuenta destino(debe, haber)
        s_Sentencia = "UPDATE tbltmp_cuentas"
        s_Sentencia = s_Sentencia & " SET codcta_ajd_deb='', codcta_dst_deb="
        sSentencia = "(SELECT dst.ctdacodcta FROM ctdiastp dst"
        sSentencia = sSentencia & " WHERE dst.ctdacodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND dst.ctdacodsuc='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND dst.ctdacodastp=tbltmp_cuentas.ctctcodasrf"
        s_Sentencia = s_Sentencia & sSentencia & " AND dst.ctdaflgdeb='D')"
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "UPDATE tbltmp_cuentas"
        s_Sentencia = s_Sentencia & " SET codcta_ajd_hab='', codcta_dst_hab="
        s_Sentencia = s_Sentencia & sSentencia & " AND dst.ctdaflgdeb='C')"
        pocnnMain.Execute s_Sentencia
        ' Campos de seleccion
        s_Sentencia = "codcta, detcta, tpocta, natcta, tposdo, tpoanl,"
        s_Sentencia = s_Sentencia & " codcta_dst_deb, codcta_dst_hab, tpomon, tpotcb, tpoajd,"
        s_Sentencia = s_Sentencia & " codcta_ajd_deb, codcta_ajd_hab, indajd, indcco, inddoc, indmoe, indpsp"
      ElseIf nContador = 2 Then   ' Auxiliares
        ' Genero la tabla temporal e inserto los registros
        sSentencia = " SELECT DISTINCT clmccodcli AS codaux, clmcnomcli AS razaux,"
        sSentencia = sSentencia & " clmcnomcli AS nomaux, clmcnomcli AS apepataux, clmcnomcli AS apemataux,"
        sSentencia = sSentencia & " clmccodtipc AS tipaux, clmcnroruc olrucaux,"
        sSentencia = sSentencia & " clmcnewruc AS rucaux, 'S' AS indcli, 'N' AS indprv, 'N' AS indotr"
        sSentencia = sSentencia & " FROM cldmclie"
        sSentencia = sSentencia & " WHERE clmccodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND clmccodsuc='" & cmbParametro(1).Text & "'"
        sSentencia = sSentencia & " ORDER BY codaux"
        s_Sentencia = "CREATE GLOBAL TEMPORARY TABLE tbltmp_aux AS"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        s_Sentencia = "INSERT INTO tbltmp_aux"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        
        ' Actualizo indicador de provedor
        s_Sentencia = "UPDATE tbltmp_aux"
        s_Sentencia = s_Sentencia & " SET indprv='S'"
        s_Sentencia = s_Sentencia & " WHERE EXISTS(SELECT * FROM cpdmprov prv"
        s_Sentencia = s_Sentencia & " WHERE prv.cpmpcodcia='" & cmbParametro(0).Text & "'"
        s_Sentencia = s_Sentencia & " AND prv.cpmpcodsuc='" & cmbParametro(1).Text & "'"
        s_Sentencia = s_Sentencia & " AND prv.cpmpcodprv=tbltmp_aux.codaux)"
        pocnnMain.Execute s_Sentencia
        
        ' Inserto los proveedores no existentes
        s_Sentencia = "INSERT INTO tbltmp_aux"
        sSentencia = " SELECT DISTINCT cpmpcodprv AS codaux, cpmpnomprv AS razaux,"
        sSentencia = sSentencia & " cpmpnomprv AS nomaux, cpmpnomprv AS apepataux, cpmpnomprv AS apemataux,"
        sSentencia = sSentencia & " cpmpcodtipp AS tipaux, cpmpnroruc olrucaux,"
        sSentencia = sSentencia & " cpmpnewruc AS rucaux, 'N' AS indcli, 'S' AS indprv, 'N' AS indotr"
        sSentencia = sSentencia & " FROM cpdmprov"
        sSentencia = sSentencia & " WHERE cpmpcodcia='" & cmbParametro(0).Text & "'"
        sSentencia = sSentencia & " AND cpmpcodsuc='" & cmbParametro(1).Text & "'"
        sSentencia = sSentencia & " AND NOT EXISTS(SELECT * FROM tbltmp_aux"
        sSentencia = sSentencia & " WHERE tbltmp_aux.codaux=cpdmprov.cpmpcodprv)"
        sSentencia = sSentencia & " ORDER BY codaux"
        s_Sentencia = s_Sentencia & sSentencia
        pocnnMain.Execute s_Sentencia
        ' campos de sleciion
        s_Sentencia = "CodAux, RazAux, NomAux, ApePatAux, ApeMatAux, TipAux, RucAux, IndCli, IndPrv, IndOtr"
      End If
      ' Obtengo el archivo de texto
      sArchivo = Choose(nContador + 1, "dro", "pct", "aux", "tdc", "cco", "tpc")
      sColumnas = Choose(nContador + 1, "coddro, desdro", s_Sentencia, s_Sentencia, "codtpodoc, destpodoc, abvtipdoc, sgntpodoc", "codcencos, ccm_descrip", "cntpfecdia, cntpvalcom, cntpvalvta")
      ' Genero el archivo y abro el recordset temporal
      sTabla = Choose(nContador + 1, "tmpdro", "tbltmp_cuentas", "tbltmp_aux", "tmptpodoc", "tmpcencos", "cntipcam")
      sWhere = Choose(nContador + 1, "", "", "", "", "", "cntpcodmon='" & gsTpoMon_Sgn_ME & "'")
      ' Obtengo el archivo de texto libre
      nArchivo = FreeFile
      sSentencia = " SELECT DISTINCT " & sColumnas
      sSentencia = sSentencia & " FROM " & sTabla
      sSentencia = sSentencia & IIf(sWhere <> "", " WHERE ", "") & sWhere
      sSentencia = sSentencia & " ORDER BY 1"
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
        lblProgreso(1).Caption = Choose(gsIdioma, "Exportando Archivo: ", "Exporting File: ") & Trim(chkCentraTabla(nContador).Caption)
        nNumRegistros = porstTmp.RecordCount
        pgbProgreso(1).Max = nNumRegistros
        pgbProgreso(1).Value = pgbProgreso(1).Min
        nRegistro = 0
        sArchivo = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & sArchivo & ".txt"
        ' Elimino archivo de texto si existe
        If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
        If Dir$(sArchivo, vbNormal) = "" Then
          Open sArchivo For Output Access Write Lock Read Write As #nArchivo
          While Not porstTmp.EOF
            nRegistro = nRegistro + 1
            ' Diseño y grabro la linea en el archivo
            sMilinea = ""
            Select Case nContador
             Case 0
              sMilinea = sMilinea & Trim(porstTmp!coddro) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(porstTmp!desdro) & sCaracter
             Case 1
              sMilinea = sMilinea & Trim(porstTmp!codcta) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(porstTmp!detcta) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!TpoCTA) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!NatCta) & sCaracter
              sMilinea = sMilinea & Left(IIf((porstTmp!TpoSdo >= "0" And porstTmp!TpoSdo <= "5"), TPOSDO_INV_TXT, TPOSDO_RES_TXT), 1) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!TpoAnl) & sCaracter
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcta_dst_deb), porstTmp!codcta_dst_deb, "")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcta_dst_hab), porstTmp!codcta_dst_hab, "")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!tpomon), porstTmp!tpomon, "")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!TpoTcb), porstTmp!TpoTcb, "")) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!tpoajd) & sCaracter
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_AjD_Deb), porstTmp!CodCta_AjD_Deb, "")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_AjD_Hab), porstTmp!CodCta_AjD_Hab, "")) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!IndAjD) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!indcco) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!IndDoc) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!IndMoe) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!IndPsp) & sCaracter
             Case 2
              sMilinea = sMilinea & Trim(porstTmp!codaux) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(porstTmp!razAux) & sCaracter
              If Trim(porstTmp!tipaux) = "50" Then
                n_Posicion1 = InStr(Trim(porstTmp!NomAux), " ")
                n_Posicion1 = IIf(n_Posicion1 = 0, Len(Trim(porstTmp!NomAux)), n_Posicion1)
                s_Sentencia = Left(Trim(porstTmp!NomAux), n_Posicion1 - 1)
                n_Posicion2 = InStr(n_Posicion1 + 1, Trim(porstTmp!NomAux), " ")
                n_Posicion2 = IIf(n_Posicion2 = 0, Len(Trim(porstTmp!NomAux)), n_Posicion2)
                sSentencia = Trim(Mid(Trim(porstTmp!NomAux), n_Posicion1 + 1, (n_Posicion2 - n_Posicion1)))
                sMilinea = sMilinea & Trim(Mid(porstTmp!NomAux, n_Posicion2 + 1)) & sCaracter
                sMilinea = sMilinea & Trim(s_Sentencia) & sCaracter
                sMilinea = sMilinea & Trim(sSentencia) & sCaracter
              Else
                sMilinea = sMilinea & sCaracter & sCaracter & sCaracter
              End If
              sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!rucaux), porstTmp!rucaux, "")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndCli) Or porstTmp!IndCli = INDAUX_CLI_INA, "N", "S")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndPrv) Or porstTmp!IndPrv = INDAUX_PRV_INA, "N", "S")) & sCaracter
              sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndOtr) Or porstTmp!IndOtr = INDAUX_OTR_INA, "N", "S")) & sCaracter
             Case 3
              sMilinea = sMilinea & Trim(porstTmp!codtpodoc) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(porstTmp!destpodoc) & sCaracter
              sMilinea = sMilinea & Trim(porstTmp!abvtipdoc) & sCaracter
              sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!sgntpodoc), "+", porstTmp!sgntpodoc)) & sCaracter
             Case 4
              sMilinea = sMilinea & Trim(porstTmp!codcencos) & sCaracter
              sMilinea = sMilinea & pfSacaApoRet(porstTmp!ccm_descrip) & sCaracter
             Case 5
              sMilinea = sMilinea & Format(porstTmp!cntpfecdia, "dd/mm/yyyy") & sCaracter
              sMilinea = sMilinea & Format(porstTmp!cntpvalcom, "####0.0000") & sCaracter
              sMilinea = sMilinea & Format(porstTmp!cntpvalvta, "####0.0000") & sCaracter
            End Select
            Print #nArchivo, sMilinea
            pgbProgreso(1).Value = nRegistro
            DoEvents
            porstTmp.MoveNext
          Wend
          Close #nArchivo
        End If
        If nContador = 1 Then     ' Plan de cuentas
          pocnnMain.Execute "DROP TABLE tbltmp_cuentas"
        ElseIf nContador = 2 Then ' Auxiliares
          pocnnMain.Execute "DROP TABLE tbltmp_aux"
        End If
      End If
      porstTmp.Close
    End If
    pgbProgreso(0).Value = nContador + 1
  Next nContador
  Set porstTmp = Nothing

End Sub

Private Sub ppEncabezadoRpt(toRpt As CrystalReport, tsTit As String, tsFEm As Date, tbImpMesAno As Boolean)
   Dim dnContador As Integer
   
  'Inicializa a "" todas las fórmulas, para no tener problemas de un reporte a otro. Esto es necesario por estar usando un único objeto rptMain de Crystal.
   For dnContador = 0 To 40: toRpt.Formulas(dnContador) = "": Next dnContador
   
   With toRpt
      .Formulas(0) = "mSistema='" & gsNomSis & "'"
      .Formulas(1) = "mEmpresa='" & Trim(gsRazEmp) & "'"
      .Formulas(2) = "mTitulo='" & tsTit & "'"
      .WindowTitle = tsTit
      .Formulas(3) = "mFeReporte='" & Format(tsFEm, "dddd") & "," _
                                    & Format(tsFEm, " d ") & "de" _
                                    & Format(tsFEm, " mmmm ") & "de" _
                                    & Format(tsFEm, " yyyy") _
                                    & "'"
      .Formulas(4) = "mHrReporte='" & Format(Time(), "hh:mm:ss AMPM") & "'"
      If tbImpMesAno Then .Formulas(5) = "mPeriodo='" & Format(CDate(IIf(gsMesAct = "00", "01", IIf(gsMesAct > "12", "12", gsMesAct)) & " " & gsAnoAct), "mmmm") & " " & gsAnoAct & "'"
      ' Inicializo las formulas de seleccion
      .SelectionFormula = ""
      .ParameterFields(0) = "Idioma;" & gsIdioma & ";true"
      .WindowShowCloseBtn = True
      .WindowShowPrintSetupBtn = True
      .WindowShowRefreshBtn = True
      .WindowShowSearchBtn = True
      .WindowShowZoomCtl = True
   End With
   
End Sub

Private Sub ppExportar_Tablas()
  Static sSentencia As String, sMilinea As String
  Static nContador As Integer, nArchivo As Integer, nFreeFile As Integer
  Static nRegistro As Long, nNumRegistros As Long
  Static sArchivo As String, sCaracter As String, sFreeFile As String
  Dim sTabla As String, sColumnas As String, sWhere As String
  Dim porstTmp As ADODB.Recordset, porstFreeFile As ADODB.Recordset

  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  sCaracter = "|"
  pgbProgreso(0).Max = chkTransTabla.Count
  pgbProgreso(0).Value = pgbProgreso(0).Min
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkTransTabla.Count - 1
    ' Verifico que se haya seleccionado
    If (chkTransTabla(nContador).Value = vbChecked) Then
      ' Obtengo el archivo de texto
      sArchivo = Choose(nContador + 1, "dro", "", "pct", "aux", "tdc", "cco", "tpc", "efi", "", "")
      'sColumnas = Choose(nContador + 1, "pdoano, coddro, detdro", "", "pdoano, codcta, detcta, tpocta, natcta, tposdo, tpoanl, codcta_dst_deb, codcta_dst_hab, tpomon, tpotcb, tpoajd, codcta_ajd_deb, codcta_ajd_hab, IndAjD, IndCCo, IndDoc, IndMoe, IndPsp", "a.CodAux, RazAux, NomAux, ApePatAux, ApeMatAux, RucAux, DirAux, IndCli, IndPrv, IndOtr", "CodTDc, DetTDc, AbvTDc, SgnTDc, ForImp", "pdoano, CodCCo, DetCCo", "FehTCb, ImpTCb_Cpr, ImpTCb_Vta", "pdoano, codefi, detefi, detefix, coddpe, indcnv", "", "")
      sColumnas = Choose(nContador + 1, "pdoano, coddro, detdro", "", "pdoano,codcta,detcta,detctax,tpocta,natcta,tposdo,tpoanl,codcta_dst_deb,codcta_dst_hab,codcco_dst_deb,codcco_dst_hab,tpomon,tpotcb,tpoajd,codcta_ajd_deb,codcta_ajd_hab,codcco_ajd_deb,codcco_ajd_hab,indajd, codcta_crr_deu, codcta_crr_acr, codcco_def,indcco,inddoc,indmoe,indpsp,indfjo,codbco,estcta", "a.CodAux, RazAux, NomAux, ApePatAux, ApeMatAux, RucAux, DirAux, IndCli, IndPrv, IndOtr, tpodci", "CodTDc, DetTDc, AbvTDc, SgnTDc, ForImp", "pdoano, CodCCo, DetCCo", "FehTCb, ImpTCb_Cpr, ImpTCb_Vta", "pdoano, codefi, detefi, detefix, coddpe, indcnv", "", "")
      ' Genero el archivo y abro el recordset temporal
      sTabla = Choose(nContador + 1, "CoDro a", "", "cocta a", "(TgAux a LEFT JOIN TgAuxNat b ON a.codemp=b.codemp AND a.CodAux=b.CodAux)", "tgtdc a", "cocco a", "tgtcb a", "coefi a", "", "")
      'sWhere = Choose(nContador + 1, "AND a.pdoano='" & gsAnoAct & "' ", "AND a.pdoano='" & gsAnoAct & "' ", "", "", "", "AND a.pdoano='" & gsAnoAct & "' ", "", "AND a.pdoano='" & gsAnoAct & "' ", "", "")
        
      sWhere = Choose(nContador + 1, "AND a.pdoano='" & gsAnoAct & "' ", "AND a.pdoano='" & gsAnoAct & "' ", "AND a.pdoano='" & gsAnoAct & "' ", "", "", "AND a.pdoano='" & gsAnoAct & "' ", "", "AND a.pdoano='" & gsAnoAct & "' ", "", "")
        
      If sTabla <> "" Then
        ' Obtengo el archivo de texto libre
        nArchivo = FreeFile
        
        sSentencia = "SELECT " & sColumnas & " "
        sSentencia = sSentencia & "FROM " & sTabla & " "
        sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & sWhere
        sSentencia = sSentencia & "ORDER BY 1"
        
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
          lblProgreso(1).Caption = Choose(gsIdioma, "Exportando Archivo: ", "Exporting File: ") & Trim(chkTransTabla(nContador).Caption)
          nNumRegistros = porstTmp.RecordCount
          pgbProgreso(1).Max = nNumRegistros
          pgbProgreso(1).Value = pgbProgreso(1).Min
          nRegistro = 0
          sArchivo = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & sArchivo & ".txt"
          ' Elimino archivo de texto si existe
          If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
          If Dir$(sArchivo, vbNormal) = "" Then
            Open sArchivo For Output Access Write Lock Read Write As #nArchivo
            ' Estados financieros
            If nContador = 7 Then
              ' Obtengo el archivo de texto libre
              nFreeFile = FreeFile
              sFreeFile = dlbDirectorio(tabProceso.Tab).path & "\" & gsRUCEmp & "cef.txt"
              ' Elimino archivo de texto si existe
              If Dir$(sFreeFile, vbNormal) <> "" Then Kill sFreeFile
              Open sFreeFile For Output Access Write Lock Read Write As #nFreeFile
              Set porstFreeFile = New ADODB.Recordset
            End If
            While Not porstTmp.EOF
              nRegistro = nRegistro + 1
              ' Diseño y grabro la linea en el archivo
              sMilinea = ""
              Select Case nContador
               Case 0
                sMilinea = sMilinea & Trim(porstTmp!pdoano) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!coddro) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!DetDro) & sCaracter
               Case 2
                'sMilinea = sMilinea & Trim(porstTmp!pdoano) & sCaracter
                'sMilinea = sMilinea & Trim(porstTmp!codcta) & sCaracter
                'sMilinea = sMilinea & Trim(porstTmp!detcta) & sCaracter
                'sMilinea = sMilinea & Left(IIf(porstTmp!TpoCTA = TPOCTA_TRA, TPOCTA_TRA_TXT, TPOCTA_TIT_TXT), 1) & sCaracter
                'sMilinea = sMilinea & Left(IIf(porstTmp!NatCta = NATCTA_DEU, NATCTA_DEU_TXT, NATCTA_ACR_TXT), 1) & sCaracter
                'sMilinea = sMilinea & Left(IIf(porstTmp!TpoSdo = TPOSDO_INV, TPOSDO_INV_TXT, IIf(porstTmp!TpoSdo = TPOSDO_RES, TPOSDO_RES_TXT, IIf(porstTmp!TpoSdo = TPOSDO_FUN, TPOSDO_FUN_TXT, IIf(porstTmp!TpoSdo = TPOSDO_NAT, TPOSDO_NAT_TXT, "A")))), 1) & sCaracter
                'sMilinea = sMilinea & Left(IIf(porstTmp!TpoAnl = TPOANL_CTA, TPOANL_CTA_TXT, IIf(porstTmp!TpoAnl = TPOANL_AUX, "A", TPOANL_DOC_TXT)), 1) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_Dst_Deb), porstTmp!CodCta_Dst_Deb, "")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_Dst_Hab), porstTmp!CodCta_Dst_Hab, "")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!tpomon), porstTmp!tpomon, "")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!TpoTcb), porstTmp!TpoTcb, "")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!tpoajd) Or porstTmp!tpoajd = INDAJD_INA, "N", "S")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_AjD_Deb), porstTmp!CodCta_AjD_Deb, "")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_AjD_Hab), porstTmp!CodCta_AjD_Hab, "")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndAjD) Or porstTmp!IndAjD = INDAJD_INA, "N", "S")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndCCo) Or porstTmp!IndCCo = INDCCO_INA, "N", "S")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndDoc) Or porstTmp!IndDoc = INDDOC_INA, "N", "S")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndMoe) Or porstTmp!IndMoe = INDMOE_INA, "N", "S")) & sCaracter
                'sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndPsp) Or porstTmp!IndPsp = INDPSP_INA, "N", "S")) & sCaracter
                
                sMilinea = sMilinea & Trim(porstTmp!pdoano) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!codcta) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!detcta) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!DetCtax) & sCaracter
                sMilinea = sMilinea & Left(IIf(porstTmp!TpoCTA = TPOCTA_TRA, TPOCTA_TRA_TXT, TPOCTA_TIT_TXT), 1) & sCaracter
                sMilinea = sMilinea & Left(IIf(porstTmp!NatCta = NATCTA_DEU, NATCTA_DEU_TXT, NATCTA_ACR_TXT), 1) & sCaracter
                sMilinea = sMilinea & Left(IIf(porstTmp!TpoSdo = TPOSDO_INV, TPOSDO_INV_TXT, IIf(porstTmp!TpoSdo = TPOSDO_RES, TPOSDO_RES_TXT, IIf(porstTmp!TpoSdo = TPOSDO_FUN, TPOSDO_FUN_TXT, IIf(porstTmp!TpoSdo = TPOSDO_NAT, TPOSDO_NAT_TXT, "A")))), 1) & sCaracter
                sMilinea = sMilinea & Left(IIf(porstTmp!TpoAnl = TPOANL_CTA, TPOANL_CTA_TXT, IIf(porstTmp!TpoAnl = TPOANL_AUX, "A", TPOANL_DOC_TXT)), 1) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcta_dst_deb), porstTmp!codcta_dst_deb, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcta_dst_hab), porstTmp!codcta_dst_hab, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcco_dst_deb), porstTmp!codcco_dst_deb, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcco_dst_hab), porstTmp!codcco_dst_hab, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!tpomon), porstTmp!tpomon, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!TpoTcb), porstTmp!TpoTcb, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!tpoajd) Or porstTmp!tpoajd = INDAJD_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_AjD_Deb), porstTmp!CodCta_AjD_Deb, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCta_AjD_Hab), porstTmp!CodCta_AjD_Hab, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCCo_AjD_Deb), porstTmp!CodCCo_AjD_Deb, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!CodCCo_AjD_Hab), porstTmp!CodCCo_AjD_Hab, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndAjD) Or porstTmp!IndAjD = INDAJD_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcta_crr_deu), porstTmp!codcta_crr_deu, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codcco_def), porstTmp!codcco_def, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!indcco) Or porstTmp!indcco = INDCCO_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndDoc) Or porstTmp!IndDoc = INDDOC_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndMoe) Or porstTmp!IndMoe = INDMOE_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndPsp) Or porstTmp!IndPsp = INDPSP_INA, "N", "S")) & sCaracter
                'OJO
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndFjo) Or porstTmp!IndFjo = 1, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!codbco), porstTmp!codbco, "")) & sCaracter
                sMilinea = sMilinea & IIf(porstTmp!EstCta = "A", "A", "I") & sCaracter
                
               Case 3
                sMilinea = sMilinea & Trim(porstTmp!codaux) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!razAux) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!NomAux), porstTmp!NomAux, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!ApePatAux), porstTmp!ApePatAux, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!ApeMatAux), porstTmp!ApeMatAux, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!rucaux), porstTmp!rucaux, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!DirAux), porstTmp!DirAux, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndCli) Or porstTmp!IndCli = INDAUX_CLI_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndPrv) Or porstTmp!IndPrv = INDAUX_PRV_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!IndOtr) Or porstTmp!IndOtr = INDAUX_OTR_INA, "N", "S")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!TpoDci), porstTmp!TpoDci, "00")) & sCaracter
               Case 4
                sMilinea = sMilinea & Trim(porstTmp!codtdc) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!dettdc) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!AbvTDc) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!SgnTDc) Or porstTmp!SgnTDc = SGNTDC_POS, "+", "-")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(IsNull(porstTmp!forimp), "0", porstTmp!forimp)) & sCaracter
               Case 5
                sMilinea = sMilinea & Trim(porstTmp!pdoano) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!codcco) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!detcco) & sCaracter
               Case 6
                sMilinea = sMilinea & Format(porstTmp!FehTCb, "dd/mm/yyyy") & sCaracter
                sMilinea = sMilinea & Format(porstTmp!ImpTCb_Cpr, "####0.0000") & sCaracter
                sMilinea = sMilinea & Format(porstTmp!ImpTCb_Vta, "####0.0000") & sCaracter
               Case 7
                sMilinea = sMilinea & Trim(porstTmp!pdoano) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!CodEfi) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!DetEFi) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!DetEFix), porstTmp!DetEFix, "")) & sCaracter
                sMilinea = sMilinea & Trim(IIf(Not IsNull(porstTmp!coddpe), porstTmp!coddpe, "")) & sCaracter
                sMilinea = sMilinea & Trim(porstTmp!IndCnv) & sCaracter
                
                ' Detalle de estado financiero
                sSentencia = "SELECT a.pdoano, a.codefi, a.nrolin, a.detlin, a.detlinx, a.tpolin, a.fmllin, a.bsepct, a.grppct, a.imp1, a.pct1, "
                sSentencia = sSentencia & "a.imp2, a.pct2, a.indlat, a.indbdesup, a.indbdeinf, a.indfondet, a.indfondet_syd, a.indfonimp "
                sSentencia = sSentencia & "FROM coefilin a "
                sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND a.pdoano='" & porstTmp!pdoano & "'"
                sSentencia = sSentencia & "AND a.codefi='" & porstTmp!CodEfi & "' "
                sSentencia = sSentencia & "ORDER BY 1"
                With porstFreeFile
                  If .State = adStateOpen Then .Close
                  .ActiveConnection = pocnnMain
                  .Source = sSentencia
                  .CursorType = adOpenDynamic
                  .LockType = adLockReadOnly
                  .Open
                End With
                While Not porstFreeFile.EOF
                  sSentencia = ""
                  sSentencia = sSentencia & Trim(porstFreeFile!pdoano) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!CodEfi) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!NroLin) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!DetLin) & sCaracter
                  sSentencia = sSentencia & Trim(IIf(Not IsNull(porstFreeFile!DetLinx), porstFreeFile!DetLinx, "")) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!TpoLin) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!FmlLin) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!BsePct) & sCaracter
                  sSentencia = sSentencia & Trim(IIf(Not IsNull(porstFreeFile!grppct), porstFreeFile!grppct, "")) & sCaracter
                  sSentencia = sSentencia & CDec(porstFreeFile!imp1) & sCaracter
                  sSentencia = sSentencia & CDec(porstFreeFile!pct1) & sCaracter
                  sSentencia = sSentencia & CDec(porstFreeFile!imp2) & sCaracter
                  sSentencia = sSentencia & CDec(porstFreeFile!pct2) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!IndLat) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!IndBdeSup) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!IndBdeInf) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!IndFonDet) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!IndFonDet_Syd) & sCaracter
                  sSentencia = sSentencia & Trim(porstFreeFile!IndFonImp) & sCaracter
                  Print #nFreeFile, sSentencia
                  porstFreeFile.MoveNext
                Wend
                porstFreeFile.Close
              End Select
              Print #nArchivo, sMilinea
              pgbProgreso(1).Value = nRegistro
              DoEvents
              porstTmp.MoveNext
            Wend
            If nContador = 7 Then Close #FreeFile
            Close #nArchivo
          End If
        End If
        porstTmp.Close
      End If
    End If
    pgbProgreso(0).Value = nContador + 1
  Next nContador
  Set porstTmp = Nothing
  Set porstFreeFile = Nothing

End Sub

Private Sub ppTransfir_Proceso()
    Static sSentencia As String, sTabla As String
    Static nContador As Integer
    Static nNumRegistros As Double

    pgbProgreso(1).Max = chkImporProceso.Count
    pgbProgreso(1).Value = pgbProgreso(1).Min
    For nContador = 0 To chkImporProceso.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporProceso(nContador).Value Then
        sTabla = Choose(nContador + 1, "CoCprDoc", "CoVtaDoc", "CoHprDoc", "CoCpbDet")
        lblProgreso(1).Caption = Choose(gsIdioma, "Procesando Información: ", "Processing Information: ") & Trim(chkImporProceso(nContador).Caption)
        Select Case nContador
          Case 0
            ' Elimino la información del periodo
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = "DELETE FROM " & sTabla & " "
              sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND MesPvs='" & gsMesAct & "'"
              pocnnMain.Execute sSentencia, nNumRegistros
            End If
            ' Actualizo tabla de documentos de compras
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
            sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc "
            sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
            sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc) "
            sSentencia = sSentencia & "ORDER BY a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
            ' Actualizo tabla de cuentas documentos de compras
            If aValidar(nContador, 1) = NvlUsr_Sup Then
              pocnnMain.Execute sSentencia, nNumRegistros
              sTabla = "cocprdoccta"
              sSentencia = "INSERT INTO " & sTabla & " "
              sSentencia = sSentencia & "SELECT DISTINCT a.* "
              sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
              sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
              sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
              sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
              sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
              sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc "
              sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
              sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc "
              sSentencia = sSentencia & "AND b.TpoCnc=a.TpoCnc "
              sSentencia = sSentencia & "AND b.Orden=a.Orden) "
              sSentencia = sSentencia & "ORDER BY a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden"
              ' Actualizo tabla centro de costos cuentas documentos de compras
              If aValidar(nContador, 2) = NvlUsr_Sup Then
                pocnnMain.Execute sSentencia, nNumRegistros
                sTabla = "cocprdoccco"
                sSentencia = "INSERT INTO " & sTabla & " "
                sSentencia = sSentencia & "SELECT DISTINCT a.* "
                sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
                sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
                sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
                sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
                sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
                sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc "
                sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
                sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc "
                sSentencia = sSentencia & "AND b.TpoCnc=a.TpoCnc "
                sSentencia = sSentencia & "AND b.Orden=a.Orden "
                sSentencia = sSentencia & "AND b.CodCco=a.CodCco) "
                sSentencia = sSentencia & "ORDER BY a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCco"
              End If
            End If
          Case 1
            ' Elimino la información del periodo
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = "DELETE FROM " & sTabla & " "
              sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND MesPvs='" & gsMesAct & "'"
              pocnnMain.Execute sSentencia, nNumRegistros
            End If
            ' Actualizo tabla de documentos de ventas
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc "
            sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
            sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc) "
            sSentencia = sSentencia & "ORDER BY a.CodTDc, a.SerDoc, a.NroDoc"
            ' Actualizo tabla de cuentas documentos de ventas
            If aValidar(nContador, 1) = NvlUsr_Sup Then
              pocnnMain.Execute sSentencia, nNumRegistros
              sTabla = "covtadoccta"
              sSentencia = "INSERT INTO " & sTabla & " "
              sSentencia = sSentencia & "SELECT DISTINCT a.* "
              sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
              sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
              sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
              sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
              sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc "
              sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
              sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc "
              sSentencia = sSentencia & "AND b.TpoCnc=a.TpoCnc "
              sSentencia = sSentencia & "AND b.Orden=a.Orden) "
              sSentencia = sSentencia & "ORDER BY a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden"
              ' Actualizo tabla centro de costos cuentas documentos de ventas
              If aValidar(nContador, 2) = NvlUsr_Sup Then
                pocnnMain.Execute sSentencia, nNumRegistros
                sTabla = "covtadoccco"
                sSentencia = "INSERT INTO " & sTabla & " "
                sSentencia = sSentencia & "SELECT DISTINCT a.* "
                sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
                sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
                sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
                sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
                sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc "
                sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
                sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc "
                sSentencia = sSentencia & "AND b.TpoCnc=a.TpoCnc "
                sSentencia = sSentencia & "AND b.Orden=a.Orden "
                sSentencia = sSentencia & "AND b.CodCco=a.CodCco) "
                sSentencia = sSentencia & "ORDER BY a.CodTDc, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCco"
              End If
            End If
          Case 2
            ' Elimino la información del periodo
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = "DELETE FROM " & sTabla & " "
              sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND MesPvs='" & gsMesAct & "'"
              pocnnMain.Execute sSentencia, nNumRegistros
            End If
            ' Actualizo tabla de documentos de honorarios
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
            sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
            sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc) "
            sSentencia = sSentencia & "ORDER BY a.CodAux, a.SerDoc, a.NroDoc"
            ' Actualizo tabla de cuentas documentos de honorarios
            If aValidar(nContador, 1) = NvlUsr_Sup Then
              pocnnMain.Execute sSentencia, nNumRegistros
              sTabla = "cohprdoccta"
              sSentencia = "INSERT INTO " & sTabla & " "
              sSentencia = sSentencia & "SELECT DISTINCT a.* "
              sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
              sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
              sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
              sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
              sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
              sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
              sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc "
              sSentencia = sSentencia & "AND b.TpoCnc=a.TpoCnc "
              sSentencia = sSentencia & "AND b.Orden=a.Orden) "
              sSentencia = sSentencia & "ORDER BY a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden"
              ' Actualizo tabla centro de costos cuentas documentos de honorarios
              If aValidar(nContador, 2) = NvlUsr_Sup Then
                pocnnMain.Execute sSentencia, nNumRegistros
                sTabla = "cohprdoccco"
                sSentencia = "INSERT INTO " & sTabla & " "
                sSentencia = sSentencia & "SELECT DISTINCT a.* "
                sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
                sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
                sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
                sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
                sSentencia = sSentencia & "AND b.CodAux=a.CodAux "
                sSentencia = sSentencia & "AND b.SerDoc=a.SerDoc "
                sSentencia = sSentencia & "AND b.NroDoc=a.NroDoc "
                sSentencia = sSentencia & "AND b.TpoCnc=a.TpoCnc "
                sSentencia = sSentencia & "AND b.Orden=a.Orden "
                sSentencia = sSentencia & "AND b.CodCco=a.CodCco) "
                sSentencia = sSentencia & "ORDER BY a.CodAux, a.SerDoc, a.NroDoc, a.TpoCnc, a.Orden, a.CodCco"
              End If
            End If
          Case 3
            ' Elimino la información del periodo
            If chkEliminar(nContador).Value = vbChecked Then
              sSentencia = "DELETE FROM CoCpbCab "
              sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
              sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
              sSentencia = sSentencia & "AND mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "'"
              pocnnMain.Execute sSentencia, nNumRegistros
            End If
            ' Actualizo la tabla de cabecera de comprobantes
            sSentencia = "INSERT INTO CoCpbCab (codemp, pdoano, MesPvs, CodDro, NroCpb, FehCpb, GloCpb, GloCpbx, TpoGnr, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf) "
            sSentencia = sSentencia & "SELECT DISTINCT a.codemp, a.pdoano, a.MesPvs, a.CodDro, a.NroCpb, a.FehOpe, a.GloCpb, a.GloCpbx, a.TpoGnr, " & INDNCU_FAL & ", " & INDANU_FAL & ",  "
            sSentencia = sSentencia & "a.UsrCre, a.FyHCre, a.UsrMdf, a.FyHMdf "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND a.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM CoCpbCab b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.MesPvs=a.MesPvs "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro "
            sSentencia = sSentencia & "AND b.NroCpb=a.NroCpb) "
            sSentencia = sSentencia & " GROUP BY a.MesPvs, a.CodDro, a.NroCpb "
            sSentencia = sSentencia & "ORDER BY a.MesPvs, a.CodDro, a.NroCpb"
            pocnnMain.Execute sSentencia, nNumRegistros
            ' Actualizo la tabla de detalle de comprobantes
            sSentencia = "INSERT INTO " & sTabla & " (codemp, pdoano, MesPvs, CodDro, NroCpb, NroIte, BlqIte, FehOpe, CodCta, CodCCo, CodAux, CodTDc, SerDoc, NroDoc, FeEDoc, FeVDoc, "
            sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, TpoCtb, TpoPvs, TpoMon, TpoTcb, ImpTCb, ImpMN, ImpME, TpoGnr, Tpodoc, UsrCre, FyHCre, UsrMdf, FyHMdf) "
            sSentencia = sSentencia & "SELECT DISTINCT a.codemp, a.pdoano, a.MesPvs, a.CodDro, a.NroCpb, a.NroIte, a.BlqIte, a.FehOpe, a.CodCta, a.CodCCo, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.FeEDoc, a.FeVDoc, "
            sSentencia = sSentencia & "a.FeRDoc, a.RefDoc, a.GloIte, a.GloItex, a.TpoCtb, a.TpoPvs, a.TpoMon, a.TpoTcb, a.ImpTCb, a.ImpMN, a.ImpME, a.TpoGnr, Tpodoc,a.UsrCre, a.FyHCre, a.UsrMdf, a.FyHMdf "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND a.mespvs" & IIf(chkPeriodos.Value = vbChecked, "<='", "='") & gsMesAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.MesPvs=a.MesPvs "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro "
            sSentencia = sSentencia & "AND b.NroCpb=a.NroCpb "
            sSentencia = sSentencia & "AND b.NroIte=a.NroIte) "
            sSentencia = sSentencia & "ORDER BY a.MesPvs, a.CodDro, a.NroCpb, a.NroIte"
        End Select
        pocnnMain.Execute sSentencia, nNumRegistros
      End If
      pgbProgreso(1).Value = nContador + 1
      DoEvents
     Next nContador

End Sub

Private Sub ppTransfir_Tablas()
    Static sSentencia As String, sTabla As String
    Static nContador As Integer
    Static nNumRegistros As Double

    pgbProgreso(1).Max = chkImporTabla.Count
    pgbProgreso(1).Value = pgbProgreso(1).Min
    For nContador = 0 To chkImporTabla.Count - 1
      ' Verifico que se haya seleccionado
      If chkImporTabla(nContador).Value Then
        sTabla = Choose(nContador + 1, "CoDro", "CoCta", "TgAux", "TgTDc", "CoCCo", "TgTcb")
        lblProgreso(1).Caption = Choose(gsIdioma, "Procesando Información: ", "Processing Information: ") & Trim(chkImporTabla(nContador).Caption)
        Select Case nContador
          Case 0
            ' Actualizo tabla de diarios
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodDro=a.CodDro) "
            sSentencia = sSentencia & "ORDER BY a.CodDro"
          Case 1
            ' Actualizo tabla de plan de cuentas
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodCta=a.CodCta) "
            sSentencia = sSentencia & "ORDER BY a.CodCta"
            ' Mofico la información existente
            If chkModificar(1).Value = vbChecked Then
              pocnnMain.Execute sSentencia, nNumRegistros
              If ps_Plataforma = pSrvMySql Then
                sSentencia = "UPDATE " & sTabla & " cta, tmp" & sTabla & " tmp SET "
                sSentencia = sSentencia & "cta.DetCta = tmp.DetCta, "
                sSentencia = sSentencia & "cta.DetCtax = tmp.DetCtax, "
                sSentencia = sSentencia & "cta.TpoCta = tmp.TpoCta, "
                sSentencia = sSentencia & "cta.NatCta = tmp.NatCta, "
                sSentencia = sSentencia & "cta.TpoSdo = tmp.TpoSdo, "
                sSentencia = sSentencia & "cta.TpoAnl = tmp.TpoAnl, "
                sSentencia = sSentencia & "cta.CodCta_Dst_Deb = tmp.CodCta_Dst_Deb, "
                sSentencia = sSentencia & "cta.CodCta_Dst_Hab = tmp.CodCta_Dst_Hab, "
                sSentencia = sSentencia & "cta.codcco_dst_deb = tmp.codcco_dst_deb, "
                sSentencia = sSentencia & "cta.codcco_dst_hab = tmp.codcco_dst_hab, "
                sSentencia = sSentencia & "cta.TpoMon = tmp.TpoMon, "
                sSentencia = sSentencia & "cta.TpoTcb = tmp.TpoTcb, "
                sSentencia = sSentencia & "cta.TpoAjd = tmp.TpoAjd, "
                sSentencia = sSentencia & "cta.CodCta_AjD_Deb = tmp.CodCta_AjD_Deb, "
                sSentencia = sSentencia & "cta.CodCta_AjD_Hab = tmp.CodCta_AjD_Hab, "
                sSentencia = sSentencia & "cta.codcco_ajd_deb = tmp.codcco_ajd_deb, "
                sSentencia = sSentencia & "cta.codcco_ajd_hab = tmp.codcco_ajd_hab, "
                sSentencia = sSentencia & "cta.IndAjD = tmp.IndAjD, "
                sSentencia = sSentencia & "cta.codcta_crr_deu = tmp.codcta_crr_deu, "
                sSentencia = sSentencia & "cta.codcta_crr_acr = tmp.codcta_crr_acr, "
                sSentencia = sSentencia & "cta.codcco_def = tmp.codcco_def, "
                sSentencia = sSentencia & "cta.IndCCo = tmp.IndCCo, "
                sSentencia = sSentencia & "cta.codcco_def = tmp.codcco_def, "
                sSentencia = sSentencia & "cta.IndDoc = tmp.IndDoc, "
                sSentencia = sSentencia & "cta.IndMoe = tmp.IndMoe, "
                sSentencia = sSentencia & "cta.IndPsp = tmp.IndPsp, "
                sSentencia = sSentencia & "cta.Indfjo = tmp.Indfjo, "
                'sSentencia = sSentencia & "cta.codbco = tmp.codbco, "
                sSentencia = sSentencia & "cta.EstCta = tmp.EstCta "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND tmp.codemp=cta.codemp "
                sSentencia = sSentencia & "AND tmp.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND tmp.CodCta=cta.CodCta "
                sSentencia = sSentencia & "AND IFNULL(cta.CodCta, '')<>''"
              ElseIf ps_Plataforma = pSrvSql Then
                sSentencia = "UPDATE " & sTabla & " "
                sSentencia = sSentencia & "SET "
                sSentencia = sSentencia & "DetCta = tmp.DetCta, "
                sSentencia = sSentencia & "TpoCta = tmp.TpoCta, "
                sSentencia = sSentencia & "NatCta = tmp.NatCta, "
                sSentencia = sSentencia & "TpoSdo = tmp.TpoSdo, "
                sSentencia = sSentencia & "TpoAnl = tmp.TpoAnl, "
                sSentencia = sSentencia & "CodCta_Dst_Deb = tmp.CodCta_Dst_Deb, "
                sSentencia = sSentencia & "CodCta_Dst_Hab = tmp.CodCta_Dst_Hab, "
                sSentencia = sSentencia & "codcco_dst_deb = tmp.codcco_dst_deb, "
                sSentencia = sSentencia & "codcco_dst_hab = tmp.codcco_dst_hab, "
                sSentencia = sSentencia & "TpoMon = tmp.TpoMon, "
                sSentencia = sSentencia & "TpoTcb = tmp.TpoTcb, "
                sSentencia = sSentencia & "TpoAjd = tmp.TpoAjd, "
                sSentencia = sSentencia & "CodCta_AjD_Deb = tmp.CodCta_AjD_Deb, "
                sSentencia = sSentencia & "CodCta_AjD_Hab = tmp.CodCta_AjD_Hab, "
                sSentencia = sSentencia & "codcco_ajd_deb = tmp.codcco_ajd_deb, "
                sSentencia = sSentencia & "codcco_ajd_hab = tmp.codcco_ajd_hab, "
                sSentencia = sSentencia & "IndAjD = tmp.IndAjD, "
                sSentencia = sSentencia & "IndCCo = tmp.IndCCo, "
                sSentencia = sSentencia & "codcco_def = tmp.codcco_def, "
                sSentencia = sSentencia & "IndDoc = tmp.IndDoc, "
                sSentencia = sSentencia & "IndMoe = tmp.IndMoe, "
                sSentencia = sSentencia & "IndPsp = tmp.IndPsp, "
                sSentencia = sSentencia & "EstCta = tmp.EstCta "
                sSentencia = sSentencia & "FROM " & sTabla & " cta, " & ps_Prefijo & "tmp" & sTabla & " tmp "
                sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND cta.pdoano='" & gsAnoAct & "' "
                sSentencia = sSentencia & "AND tmp.codemp=cta.codemp "
                sSentencia = sSentencia & "AND tmp.pdoano=cta.pdoano "
                sSentencia = sSentencia & "AND tmp.CodCta=cta.CodCta "
                sSentencia = sSentencia & "AND ISNULL(cta.CodCta, '')<>''"
              End If
            End If
          Case 2
            ' Actualizo tabla de auxiliares en general
            sSentencia = "INSERT INTO " & sTabla & " (codemp, CodAux, RazAux, RUCAux, TpoDci, DirAux, rubro, IndCli, IndPrv, IndOtr, TpoPer, EstAux, UsrCre, FyHCre, UsrMdf, FyHMdf) "
            sSentencia = sSentencia & "SELECT DISTINCT a.codemp, a.CodAux, a.RazAux, a.RUCAux, a.TpoDci, a.DirAux, a.rubro, a.IndCli, a.IndPrv, a.IndOtr, "
            sSentencia = sSentencia & "a.TpoPer, a.EstAux, a.UsrCre, a.FyHCre, a.UsrMdf, a.FyHMdf "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux) "
            sSentencia = sSentencia & "ORDER BY a.CodAux"
            pocnnMain.Execute sSentencia, nNumRegistros
            ' Auxiliares que son personas naturales
            sSentencia = "INSERT INTO " & sTabla & "nat "
            sSentencia = sSentencia & "SELECT DISTINCT a.codemp, a.CodAux, a.NomAux, a.ApePatAux, a.ApeMatAux,null ,null,a.UsrCre, a.FyHCre, a.UsrMdf, a.FyHMdf "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.TpoPer='" & TPOPER_NAT & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & "nat b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodAux=a.CodAux) "
            sSentencia = sSentencia & "ORDER BY a.CodAux"
            ' Mofico la información existente
            If chkModificar(2).Value = vbChecked Then
              pocnnMain.Execute sSentencia, nNumRegistros
              If ps_Plataforma = pSrvMySql Then
                sSentencia = "UPDATE " & sTabla & " aux, tmp" & sTabla & " tmp SET "
                sSentencia = sSentencia & "aux.RazAux = tmp.RazAux, "
                sSentencia = sSentencia & "aux.RUCAux = tmp.RUCAux, "
                sSentencia = sSentencia & "aux.TpoDci = tmp.TpoDci, "
                sSentencia = sSentencia & "aux.DirAux = tmp.DirAux, "
                sSentencia = sSentencia & "aux.IndCli = tmp.IndCli, "
                sSentencia = sSentencia & "aux.IndPrv = tmp.IndPrv, "
                sSentencia = sSentencia & "aux.IndOtr = tmp.IndOtr, "
                sSentencia = sSentencia & "aux.TpoPer = tmp.TpoPer, "
                sSentencia = sSentencia & "aux.EstAux = tmp.EstAux "
                sSentencia = sSentencia & "WHERE aux.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.codemp=aux.codemp "
                sSentencia = sSentencia & "AND tmp.CodAux=aux.CodAux "
                sSentencia = sSentencia & "AND IFNULL(aux.CodAux, '')<>''"
                pocnnMain.Execute sSentencia, nNumRegistros
                    
                sSentencia = "UPDATE " & sTabla & "Nat aux, tmp" & sTabla & " tmp SET "
                sSentencia = sSentencia & "aux.NomAux = tmp.NomAux, "
                sSentencia = sSentencia & "aux.ApePatAux = tmp.ApePatAux, "
                sSentencia = sSentencia & "aux.ApeMatAux = tmp.ApeMatAux "
                sSentencia = sSentencia & "WHERE aux.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.TpoPer = '" & TPOPER_NAT & "' "
                sSentencia = sSentencia & "AND tmp.codemp=aux.codemp "
                sSentencia = sSentencia & "AND tmp.CodAux=aux.CodAux "
                sSentencia = sSentencia & "AND IFNULL(aux.CodAux, '')<>''"
              ElseIf ps_Plataforma = pSrvSql Then
                sSentencia = "UPDATE " & sTabla & " SET "
                sSentencia = sSentencia & "RazAux = tmp.RazAux, "
                sSentencia = sSentencia & "RUCAux = tmp.RUCAux, "
                sSentencia = sSentencia & "TpoDci = tmp.TpoDci, "
                sSentencia = sSentencia & "DirAux = tmp.DirAux, "
                sSentencia = sSentencia & "IndCli = tmp.IndCli, "
                sSentencia = sSentencia & "IndPrv = tmp.IndPrv, "
                sSentencia = sSentencia & "IndOtr = tmp.IndOtr, "
                sSentencia = sSentencia & "TpoPer = tmp.TpoPer, "
                sSentencia = sSentencia & "EstAux = tmp.EstAux "
                sSentencia = sSentencia & "FROM " & sTabla & " aux, " & ps_Prefijo & "tmp" & sTabla & " tmp "
                sSentencia = sSentencia & "WHERE aux.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND tmp.codemp=aux.codemp "
                sSentencia = sSentencia & "AND tmp.CodAux=aux.CodAux "
                sSentencia = sSentencia & "AND ISNULL(aux.CodAux, '')<>''"
                pocnnMain.Execute sSentencia, nNumRegistros
                    
                sSentencia = "UPDATE " & sTabla & "Nat SET "
                sSentencia = sSentencia & "NomAux = tmp.NomAux, "
                sSentencia = sSentencia & "ApePatAux = tmp.ApePatAux, "
                sSentencia = sSentencia & "ApeMatAux = tmp.ApeMatAux "
                sSentencia = sSentencia & "FROM " & sTabla & "Nat aux, " & ps_Prefijo & "tmp" & sTabla & " tmp "
                sSentencia = sSentencia & "WHERE aux.codemp='" & gsCodEmp & "' "
                sSentencia = sSentencia & "AND aux.TpoPer = '" & TPOPER_NAT & "' "
                sSentencia = sSentencia & "AND tmp.codemp=aux.codemp "
                sSentencia = sSentencia & "AND tmp.CodAux=aux.CodAux "
                sSentencia = sSentencia & "AND ISNULL(aux.CodAux, '')<>''"
              End If
            End If
          Case 3
            ' Actualizo tabla de tipo de documento
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.CodTDc=a.CodTDc) "
            sSentencia = sSentencia & "ORDER BY a.CodTDc"
          Case 4
            ' Actualizo tabla de centro de costo
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
            sSentencia = sSentencia & "AND b.CodCCo=a.CodCCo) "
            sSentencia = sSentencia & "ORDER BY a.CodCCo"
          Case 5
            ' Actualizo la tabla de tipo de cambio
            sSentencia = "INSERT INTO " & sTabla & " "
            sSentencia = sSentencia & "SELECT DISTINCT a.* "
            sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmp" & sTabla & " a "
            sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
            sSentencia = sSentencia & "AND NOT EXISTS (SELECT * FROM " & sTabla & " b "
            sSentencia = sSentencia & "WHERE b.codemp=a.codemp "
            sSentencia = sSentencia & "AND b.FehTCb=a.FehTCb) "
            sSentencia = sSentencia & "ORDER BY a.FehTCb"
        End Select
        pocnnMain.Execute sSentencia, nNumRegistros
      End If
      pgbProgreso(1).Value = nContador + 1
      DoEvents
     Next nContador

End Sub

Private Sub pRegistro_Texto(ByVal sLinea As String, ByVal nColumnas As Integer, ByRef aRegistros)
    Static nCampo As Integer
    Static nInicio As Integer, nLongitud As Integer
    ReDim aRegistros(nColumnas)
    nInicio = 1
    For nCampo = 1 To nColumnas
      nLongitud = Abs(InStr(nInicio, sLinea, "|") - nInicio)
      aRegistros(nCampo) = Mid$(sLinea, nInicio, nLongitud)
      nInicio = nInicio + (nLongitud + 1)
    Next nCampo

End Sub

Private Sub cmdTablas_Click(Index As Integer)
  
  If Index = 0 Then
    frmMeDroGrd.Show vbModal
  ElseIf Index = 1 Then
    frmMeTDcGrd.Show vbModal
  Else
    frmMeCCoGrd.Show vbModal
  End If
End Sub

Private Sub dlbDirectorio_Change(Index As Integer)
  flbArchivo.path = dlbDirectorio(0).path
  flbArchivo.Refresh
End Sub
Private Sub drvUnidad_Change(Index As Integer)
  dlbDirectorio(Index).path = drvUnidad(Index).Drive
  dlbDirectorio(Index).Refresh
End Sub
Private Sub Form_Activate()
   If cmdSalir.Enabled Then cmdSalir.SetFocus
End Sub

Private Sub Form_Load()


  Dim n_Contador As Integer
  
  drvUnidad(0).Drive = gsRutSis
  dlbDirectorio(0).path = gsRutSis
  flbArchivo.path = dlbDirectorio(0).path
  flbArchivo.Pattern = gsRUCEmp & "*.txt"
  drvUnidad(1).Drive = gsRutSis
  dlbDirectorio(1).path = gsRutSis
  drvUnidad(2).Drive = gsRutSis
  dlbDirectorio(2).path = gsRutSis
  
  ' Visualización de atributos
  lblModificar.Caption = Choose(gsIdioma, "Modificar", "Modify")
  lblModificar.Visible = (gsNvlUsr = NvlUsr_Adm)
  chkModificar(1).Visible = (gsNvlUsr = NvlUsr_Adm)
  chkModificar(2).Visible = (gsNvlUsr = NvlUsr_Adm)
  
  lblEliminar.Caption = Choose(gsIdioma, "Eliminar", "Delete")
  lblEliminar.Visible = (gsNvlUsr = NvlUsr_Adm)
  chkEliminar(0).Visible = (gsNvlUsr = NvlUsr_Adm)
  chkEliminar(1).Visible = (gsNvlUsr = NvlUsr_Adm)
  chkEliminar(2).Visible = (gsNvlUsr = NvlUsr_Adm)
  chkEliminar(3).Visible = (gsNvlUsr = NvlUsr_Adm)

  For n_Contador = 1 To 10
    cmbParametro(0).AddItem Format(n_Contador, "000")
    cmbParametro(1).AddItem Format(n_Contador, "000")
  Next n_Contador
  cmbParametro(0).ListIndex = 0
  cmbParametro(1).ListIndex = 0

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(8, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Directorio :", "Archivos :", "Directorio :", "Directorio :", "Compañia :", "Sucursal :", "Empresa :", "Ejercicio :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Directory :", "Files :", "Directory :", "Directory :", "Company :", "Branch :", "Company :", "Fiscal year :")
  Next nElemento
  tabProceso.TabCaption(0) = Choose(gsIdioma, "Importación", "Import")
  frmTablas(0).Caption = Choose(gsIdioma, " Tablas ", " Tables ")
  frmTablas(1).Caption = Choose(gsIdioma, " Tablas ", " Tables ")
  frmTablas(2).Caption = Choose(gsIdioma, " Tablas ", " Tables ")
  For nElemento = 0 To chkImporTabla.Count - 1
    If gsIdioma = NvlUsr_Sup Then
      chkImporTabla(nElemento).Caption = Choose(nElemento + 1, "&Diario", "&Plan de Cuentas", "&Auxiliares", "&Tipo de Documentos", "&Centro de Costo", "Tipo de Ca&mbio")
      chkCentraTabla(nElemento).Caption = Choose(nElemento + 1, "&Diario", "&Plan de Cuentas", "&Auxiliares", "&Tipo de Documentos", "&Centro de Costo", "Tipo de Ca&mbio")
    Else
      chkImporTabla(nElemento).Caption = Choose(nElemento + 1, "&Journal", "&Plan of Account", "&Auxiliaries", "&Type of Documents", "&Cost Center", "&Rate of Exchange")
      chkCentraTabla(nElemento).Caption = Choose(nElemento + 1, "&Journal", "&Plan of Account", "&Auxiliaries", "&Type of Documents", "&Cost Center", "&Rate of Exchange")
    End If
  Next nElemento
    
  For nElemento = 0 To chkTransTabla.Count - 1
    If gsIdioma = NvlUsr_Sup Then
      chkTransTabla(nElemento).Caption = Choose(nElemento + 1, "&Diario", "Entidad &Bancaria", "&Plan de Cuentas", "&Auxiliares", "&Tipo de Documentos", "&Centro de Costo", "Tipo de Ca&mbio", "Estado &Financiero", "A&siento Tipo", "&Otras Tablas")
    Else
      chkTransTabla(nElemento).Caption = Choose(nElemento + 1, "&Journal", "Entidad &Bank", "&Plan of Account", "&Auxiliaries", "&Type of Documents", "&Cost Center", "&Rate of Exchange", "Estado &Financiero", "&Voucher Type", "&Others Tables")
    End If
  Next nElemento
  frmProceso(0).Caption = Choose(gsIdioma, " Transacciones ", " Transactions ")
  frmProceso(2).Caption = Choose(gsIdioma, " Transacciones ", " Transactions ")
  For nElemento = 0 To chkImporProceso.Count - 1
    If gsIdioma = NvlUsr_Sup Then
      chkImporProceso(nElemento).Caption = Choose(nElemento + 1, "Registro de &Compras", "Registro de &Ventas", "Registro de &Honorarios", "&Comprobantes de Diario")
      chkCentraProceso(nElemento).Caption = Choose(nElemento + 1, "Registro de &Compras", "Registro de &Ventas", "Registro de &Honorarios", "&Comprobantes de Diario")
    Else
      chkImporProceso(nElemento).Caption = Choose(nElemento + 1, "P&urchase Register", "&Sales Register", "&Feed Register", "&Journal Vouchers")
      chkCentraProceso(nElemento).Caption = Choose(nElemento + 1, "P&urchase Register", "&Sales Register", "&Feed Register", "&Journal Vouchers")
    End If
  Next nElemento
  tabProceso.TabCaption(1) = Choose(gsIdioma, "Transferencia", "Transfers")
  tabProceso.TabCaption(2) = Choose(gsIdioma, "Centralización", "Centralize")
  chkPeriodos.Caption = Choose(gsIdioma, "Multiples Periodos", "Multiple periods")
  chkCorrelativo.Caption = Choose(gsIdioma, "Enumerar Comprobantes", "Enumerate Vouchers")
  chkVerificar.Caption = Choose(gsIdioma, "Verificar Equivalente", "Verify Equivalent")
  frmUbicacion(0).Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  frmUbicacion(1).Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  frmUbicacion(2).Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  fraCentra.Caption = Choose(gsIdioma, " Parámetros ", " Parameters ")
  lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Información...", "Processing Information...")
  lblProgreso(1).Caption = Choose(gsIdioma, "Procesando Archivo...", "Processing File...")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
  frmProceso(0).Caption = Choose(gsIdioma, " Transferir Información ", " Transfer Information ")
  Set porstEmpresa = New ADODB.Recordset
  With porstEmpresa
    .ActiveConnection = CONNSTRG & gsNomBDC
    .Source = "SELECT DISTINCTROW emp.codemp, emp.razemp, emp.rucemp "
    .Source = .Source & "FROM tgemp emp, sgpms oxu "
    .Source = .Source & "WHERE oxu.codsis='CO' "
    .Source = .Source & "AND oxu.codemp=emp.codemp "
    .Source = .Source & "AND oxu.codusr='" & gsCodUsr & "' "
    .Source = .Source & "ORDER BY emp.codemp"
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  txtDato(0).MaxLength = porstEmpresa.Fields("codemp").DefinedSize
  
  ' Cargo ejerccio de trabajo y recupera información
  For n_Contador = (Val(gsAnoAct) - 3) To Val(gsAnoAct)
    cboEjercicio.AddItem Choose(gsIdioma, "Ejercicio ", "Fiscal year ") & n_Contador
  Next n_Contador
  cboEjercicio.ListIndex = 3
  
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
   Case 0                   'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub
