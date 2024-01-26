VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form fAbcCtsPeriodoSub 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "abcctsperiodsub.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7740
   Begin TabDlg.SSTab tabRegister 
      Height          =   3825
      Left            =   75
      TabIndex        =   42
      Top             =   600
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6747
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "abcctsperiodsub.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "shpCuadro(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDato(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDato(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkIngreso"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkProcesa"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmCuadro(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmCuadro(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDescripcion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbPeriodo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNumero(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNumero(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNumero(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbEjercicio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.ComboBox cmbEjercicio 
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "abcctsperiodsub.frx":0028
         Left            =   285
         List            =   "abcctsperiodsub.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1305
         Width           =   850
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   1
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1305
         Width           =   800
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   0
         Left            =   3345
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1305
         Width           =   800
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   2
         Left            =   5520
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1305
         Width           =   800
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "abcctsperiodsub.frx":002C
         Left            =   1290
         List            =   "abcctsperiodsub.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1305
         Width           =   1590
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         MaxLength       =   50
         TabIndex        =   3
         Top             =   555
         Width           =   5265
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1340
         MaxLength       =   8
         TabIndex        =   1
         Top             =   210
         Width           =   980
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1050
         Index           =   1
         Left            =   4650
         TabIndex        =   22
         Top             =   1755
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   1852
         _StockProps     =   14
         Caption         =   " Estado "
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
         Begin Threed.SSOption optEstado 
            Height          =   180
            Index           =   0
            Left            =   180
            TabIndex        =   23
            Top             =   285
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "P&endiente"
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
         Begin Threed.SSOption optEstado 
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   24
            Top             =   525
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "&Provisionado"
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
         Begin Threed.SSOption optEstado 
            Height          =   180
            Index           =   2
            Left            =   165
            TabIndex        =   25
            Top             =   765
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "&Cancelado"
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
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1620
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1770
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   7276
         _ExtentY        =   2857
         _StockProps     =   14
         Caption         =   " Fecha "
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
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   0
            Left            =   285
            TabIndex        =   15
            Top             =   510
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            Format          =   139395073
            CurrentDate     =   37515
         End
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   1
            Left            =   2370
            TabIndex        =   17
            Top             =   510
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   139395073
            CurrentDate     =   37515
         End
         Begin MSMask.MaskEdBox mskFechaCan 
            Height          =   300
            Left            =   2550
            TabIndex        =   21
            Top             =   1170
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
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
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   2
            Left            =   465
            TabIndex        =   19
            Top             =   1170
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   139395073
            CurrentDate     =   37515
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelación :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   2550
            TabIndex        =   20
            Top             =   930
            Width           =   1005
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   465
            TabIndex        =   18
            Top             =   930
            Width           =   1005
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   675
            Index           =   1
            Left            =   210
            Shape           =   4  'Rounded Rectangle
            Top             =   870
            Width           =   3705
         End
         Begin VB.Label lblDato 
            Caption         =   "Final :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   2370
            TabIndex        =   16
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Inicio :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   285
            TabIndex        =   14
            Top             =   270
            Width           =   1005
         End
      End
      Begin Threed.SSCheck chkProcesa 
         Height          =   255
         Left            =   4650
         TabIndex        =   26
         Top             =   2865
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   441
         _StockProps     =   78
         Caption         =   "Actualizar al Personal"
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
      Begin Threed.SSCheck chkIngreso 
         Height          =   255
         Left            =   4650
         TabIndex        =   27
         Top             =   3165
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   441
         _StockProps     =   78
         Caption         =   "Fecha de Ingreso(dias)"
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
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meses :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4440
         TabIndex        =   9
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Años :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3345
         TabIndex        =   7
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dias :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5520
         TabIndex        =   11
         Top             =   1005
         Width           =   795
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   0
         Left            =   3075
         Shape           =   4  'Rounded Rectangle
         Top             =   945
         Width           =   3540
      End
      Begin VB.Label lblDato 
         Caption         =   "Ejercicio / Mes Remuneración :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   4
         Top             =   1005
         Width           =   2445
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   255
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
      Begin Threed.SSCommand cmdCancel 
         Height          =   360
         Left            =   6690
         TabIndex        =   29
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
         Picture         =   "abcctsperiodsub.frx":0030
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6300
         TabIndex        =   30
         Top             =   75
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsperiodsub.frx":004C
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
         Left            =   720
         TabIndex        =   31
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   32
      Top             =   4485
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   3
         Left            =   4935
         TabIndex        =   33
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
         Picture         =   "abcctsperiodsub.frx":0068
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4545
         TabIndex        =   34
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
         Picture         =   "abcctsperiodsub.frx":0084
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2835
         TabIndex        =   35
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsperiodsub.frx":00A0
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2445
         TabIndex        =   36
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsperiodsub.frx":00BC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   3825
      Index           =   0
      Left            =   6960
      TabIndex        =   37
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6747
      _StockProps     =   15
      ForeColor       =   192
      BackColor       =   8421504
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
      BevelOuter      =   1
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   38
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edición"
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
         Outline         =   -1  'True
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   39
         Tag             =   "0"
         Top             =   810
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsperiodsub.frx":00D8
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   40
         Tag             =   "0"
         Top             =   1620
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsperiodsub.frx":00F4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   41
         Tag             =   "0"
         Top             =   2400
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsperiodsub.frx":0110
      End
   End
End
Attribute VB_Name = "fAbcCtsPeriodoSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)

End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fCtsPeriodoSub.dcaRegistro.Recordset!subcts.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fCtsPeriodoSub.dcaRegistro.Recordset!descrisub.DefinedSize
    gdl_Procedure.EditCombo "AT", cmbejercicio, 2, Me.Tag, False
    n_Index = Month(Date) - 1
    gdl_Procedure.EditCombo "AT", cmbPeriodo, n_Index, Me.Tag, False
    gdl_Procedure.EditText "AT", txtnumero(0), 0, Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtnumero(1), 0, Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtnumero(2), 0, Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(2), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditMask "AT", mskFechaCan, "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(2), False, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", chkProcesa, True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkIngreso, False, Me.Tag, True
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fCtsPeriodoSub.dcaRegistro.Recordset!subcts, Me.Tag, True, fCtsPeriodoSub.dcaRegistro.Recordset!subcts.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fCtsPeriodoSub.dcaRegistro.Recordset!descrisub), Me.Tag, False, fCtsPeriodoSub.dcaRegistro.Recordset!descrisub.DefinedSize
    n_Index = (fCtsPeriodoSub.dcaRegistro.Recordset!pdoano)
    gdl_Procedure.EditCombo "AT", cmbejercicio, (2 + (n_Index - ps_Anyo)), Me.Tag, False
    n_Index = (fCtsPeriodoSub.dcaRegistro.Recordset!pdomes)
    gdl_Procedure.EditCombo "AT", cmbPeriodo, (n_Index - 1), Me.Tag, False
    gdl_Procedure.EditText "AT", txtnumero(0), CInt(fCtsPeriodoSub.dcaRegistro.Recordset!numeroanos), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtnumero(1), CInt(fCtsPeriodoSub.dcaRegistro.Recordset!numeromeses), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtnumero(2), CInt(fCtsPeriodoSub.dcaRegistro.Recordset!numerodias), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), fCtsPeriodoSub.dcaRegistro.Recordset!fechaini, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), fCtsPeriodoSub.dcaRegistro.Recordset!fechafin, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(2), fCtsPeriodoSub.dcaRegistro.Recordset!fechaven, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditMask "AT", mskFechaCan, IIf(IsNull(fCtsPeriodoSub.dcaRegistro.Recordset!fechacan), "", fCtsPeriodoSub.dcaRegistro.Recordset!fechacan), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fCtsPeriodoSub.dcaRegistro.Recordset!estadosub = s_Estado_Ina), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fCtsPeriodoSub.dcaRegistro.Recordset!estadosub = s_Estado_Act), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(2), (fCtsPeriodoSub.dcaRegistro.Recordset!estadosub = s_Estado_Blq), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", chkProcesa, False, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", chkIngreso, False, Me.Tag, False
  End If

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_PeriodoCts As String
  
  ' Valido que el periodo no se encuentre procesado
  If Not optEstado(0).Value And Index <> 0 Then
    Beep
    MsgBox "Sub Periodo No se puede " & Choose(Index, "Eliminar", "Modificar") & " se encuentra " & IIf(optEstado(1).Value, "Provisionado", "Cancelado"), vbExclamation: Me.Tag = s_MdoData_Vis: Exit Sub
  End If
  
  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   txtDescripcion.SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim$(txtDescripcion) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim$(txtCodigo)
    s_PeriodoCts = Trim(fCtsPeriodo.dcaRegistro.Recordset!pdocts)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "pdocts", "subcts")
    a_Valores = Array(ps_ClsPlanilla, s_PeriodoCts, s_Registro)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plctsperiodosub", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fCtsPeriodoSub.dcaRegistro, fCtsPeriodoSub.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fCtsPeriodoSub.dcaRegistro.Recordset.EOF And fCtsPeriodoSub.dcaRegistro.Recordset.BOF) Or fCtsPeriodoSub.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fCtsPeriodoSub.dcaRegistro.Recordset.Find ("subcts >= '" & s_Registro & "'")
      If fCtsPeriodoSub.dcaRegistro.Recordset.EOF Then fCtsPeriodoSub.dcaRegistro.Recordset.MoveLast
    Else
      Unload Me
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  If Not l_ExistRecord Then cmdCancel_Click

End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero inicial, anterior, siguiente o final
  Select Case Index
   Case 0: fCtsPeriodoSub.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fCtsPeriodoSub.dcaRegistro.Recordset.BOF Then fCtsPeriodoSub.dcaRegistro.Recordset.MovePrevious
           If fCtsPeriodoSub.dcaRegistro.Recordset.BOF Then fCtsPeriodoSub.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fCtsPeriodoSub.dcaRegistro.Recordset.EOF Then fCtsPeriodoSub.dcaRegistro.Recordset.MoveNext
           If fCtsPeriodoSub.dcaRegistro.Recordset.EOF Then fCtsPeriodoSub.dcaRegistro.Recordset.MoveLast
   Case 3: fCtsPeriodoSub.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1, s_PeriodoCts As String
  Dim n_Annos As Integer, n_Meses As Integer
  Dim n_Dias  As Long, n_DiasIngreso As Long, n_DiasAusencia As Long
  Dim porstBusqueda As ADODB.Recordset
  Dim s_FechaIni As String, s_DesAusenciaBF As String
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If Not IsNumeric(txtnumero(0).Text) Or CInt(txtnumero(0).Text) < 0 Then Beep: MsgBox "Numero de años  no valido; verifique", vbExclamation: txtnumero(0).SetFocus: Exit Sub
  If Not IsNumeric(txtnumero(1).Text) Or CInt(txtnumero(1).Text) < 0 Or CInt(txtnumero(1).Text) > 11 Then Beep: MsgBox "Numero de meses no valido; verifique", vbExclamation: txtnumero(1).SetFocus: Exit Sub
  If Not IsNumeric(txtnumero(2).Text) Or CInt(txtnumero(2).Text) < 0 Or CInt(txtnumero(1).Text) > 29 Then Beep: MsgBox "Numero de dias no valido; verifique", vbExclamation: txtnumero(2).SetFocus: Exit Sub
  If cmbPeriodo = "" Then Beep: MsgBox "Slecciono el mes de Remuneración", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha final debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  If Not (dtpFechas(2) >= dtpFechas(1)) Then Beep: MsgBox "Fecha vencimiento debe ser mayor o igual que la fecha Final", vbExclamation: dtpFechas(2).SetFocus: Exit Sub
  If Not optEstado(2).Value And mskFechaCan.ClipText <> "" Then Beep: MsgBox "No debe Ingresar la Fecha de Cancelación", vbExclamation: mskFechaCan.SetFocus: Exit Sub
  If optEstado(2).Value And mskFechaCan.ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha de Cancelación", vbExclamation: mskFechaCan.SetFocus: Exit Sub
  If mskFechaCan.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFechaCan, 1900) Then mskFechaCan.SetFocus: Exit Sub
  End If
  If Not (Trim(cmbejercicio) & Left(cmbPeriodo, 2) >= Format(dtpFechas(0), "yyyymm") And Trim(cmbejercicio.Text) & Left(cmbPeriodo, 2) <= Format(dtpFechas(1), "yyyymm")) Then Beep: MsgBox "Mes debe ser dentro del rango de las fechas", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, s_Estado_Ina, IIf(optEstado(1).Value, s_Estado_Act, s_Estado_Blq))
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = txtCodigo
  s_PeriodoCts = Trim(fCtsPeriodo.dcaRegistro.Recordset!pdocts)
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "pdocts", "subcts", "descrisub", "pdoano", "pdomes", "numeroanos", "numeromeses", "numerodias", "fechaini", "fechafin", "fechaven", "fechacan", "estadosub", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, s_PeriodoCts, txtCodigo, Trim$(txtDescripcion), Trim(cmbejercicio.Text), Trim(Left(cmbPeriodo.Text, 2)), CInt(txtnumero(0).Text), CInt(txtnumero(1).Text), CInt(txtnumero(2).Text), Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), Format(dtpFechas(2), s_FmtFechMysql_0), Format(mskFechaCan, s_FmtFechMysql_0), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "pdocts", "subcts")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plctsperiodosub", a_Campos, a_Valores, a_Tipos) Then GoTo Error
    s_DesAusenciaBF = s_Estado_Ina
    ' Adiciono los sub-periodos en el movimiento cts por persona
    If chkProcesa.Value Then
      ' Personal fecha de ingreso dentro de rango
      s_Sql = "SELECT sub.codcls, sub.pdocts, sub.subcts, psn.codpsn, sub.pdoano, sub.pdomes, "
      s_Sql = s_Sql & "sub.numeroanos, sub.numeromeses, sub.numerodias, "
      s_Sql = s_Sql & "IF(DATE_FORMAT(sub.fechaini, '%Y%m%d')>=DATE_FORMAT(psn.fecingreso, '%Y%m%d'), sub.fechaini, psn.fecingreso) AS fechaini, "
      s_Sql = s_Sql & "sub.fechafin, sub.fechaven, sub.fechacan, sub.estadosub, psn.fecingreso, cfg.gratixasis "
      s_Sql = s_Sql & "FROM plctsperiodosub sub "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON sub.codcls=psn.codcls AND psn.ctsdeposito='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "LEFT JOIN plcfgempresa cfg ON cfg.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "WHERE sub.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND sub.pdocts='" & s_PeriodoCts & "' "
      s_Sql = s_Sql & "AND sub.subcts='" & txtCodigo.Text & "' "
      s_Sql = s_Sql & "AND sub.fechafin>=psn.fecingreso "
      s_Sql = s_Sql & "AND sub.fechafin<IFNULL(psn.fecbaja, ADDDATE(sub.fechafin, 1)) "
      s_Sql = s_Sql & "ORDER BY psn.codpsn"
      Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
        ' Creo los arreglos para la actualización
        a_Campos = Array("codcls", "pdocts", "subcts", "codpsn", "pdoano", "pdomes", "numeroanos", "numeromeses", "numerodias", "fechaini", "fechafin", "fechaven", "fechacan", "estadomov", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
        a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
        a_Where = Array("codcls", "pdocts", "subcts", "codpsn")
        While Not porstRecordset.EOF
          ' Obtengo la fecha inicial del periodo
          s_FechaIni = Format(porstRecordset!fechaini, s_FormatoFecha)
          s_FechaIni = Format(IIf(Format(porstRecordset!fecingreso, "yyyymmdd") < Format(s_FechaIni, "yyyymmdd"), DateAdd("m", 1, "01" & Mid(s_FechaIni, 3)), s_FechaIni), s_FormatoFecha)
          n_Dias = 0
          ' Obtengo los años, meses y dias
          If Not chkIngreso.Value Then
            s_Sql = "SELECT SUM(asi.diatrabajo+asi.diaprepostnatal+asi.accidente+asi.diavacaciones+asi.enfermedad+"
            s_Sql = s_Sql & "(CASE WHEN asi.codmdi_licen NOT IN('01','05','07') THEN asi.licencia ELSE 0 END)) nDias "
            s_Sql = s_Sql & "FROM plasistencia asi "
            s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON pdo.codcls=asi.codcls AND asi.codpdo=pdo.codpdo "
            s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
            s_Sql = s_Sql & "AND asi.codpsn='" & Trim(porstRecordset!codpsn) & "' "
            s_Sql = s_Sql & "AND pdo.estadopdo<>'" & s_Estado_Ina & "' "
            s_Sql = s_Sql & "AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & Format(s_FechaIni, "yyyymm") & "' "
            s_Sql = s_Sql & "AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & Format(porstRecordset!fechafin, "yyyymm") & "' "
            s_Sql = s_Sql & "GROUP BY codpsn"
            Set porstBusqueda = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
            If Not (porstBusqueda.EOF And porstBusqueda.BOF) Then
              n_Dias = CLng(porstBusqueda!nDias)
            End If
            porstBusqueda.Close
          Else
            n_Dias = gdl_Funcion.NumeroDias360Nuevo(Format(porstRecordset!fechafin, s_FormatoFecha), s_FechaIni, Format(porstRecordset!fechafin, s_FormatoFecha))
          End If
          ' Actualizo si trabajador laboral por lo menos un mes
          n_DiasIngreso = DateDiff("d", porstRecordset!fecingreso, porstRecordset!fechafin) + 1
          If ((n_DiasIngreso >= 30 Or (Format(porstRecordset!fechafin, "yyyymmdd") >= Format(DateAdd("m", 1, porstRecordset!fecingreso), "yyyymmdd"))) And n_Dias >= 30) Then
            s_DesAusenciaBF = porstRecordset!gratixasis
            ' Resta Ausencias encontradas en el período (Si el parámetro así lo indica)
            n_DiasAusencia = 0
            If s_DesAusenciaBF = s_Estado_Act Then
              n_DiasAusencia = gdl_Funcion.DiasAusenciaBS(gdl_Conexion.CadenaConexion, ps_ClsPlanilla, porstRecordset!codpsn, s_FechaIni, Format(porstRecordset!fechafin, s_FormatoFecha))
            End If
            n_Dias = n_Dias - n_DiasAusencia
            
            n_Annos = 0: n_Meses = 0
            ' Obtengo el numero de años
            If n_Dias >= 360 Then
              n_Annos = CInt(n_Dias \ 360)
              n_Dias = (n_Dias - (360 * n_Annos))
            End If
            ' Obtengo el numero de meses
            If n_Dias < 360 Then
              n_Meses = CInt(n_Dias \ 30)
              n_Dias = (n_Dias - (n_Meses * 30))
            End If
            
            ' Realizo la actualizacion de movimientos de cts
            If CInt(n_Annos + n_Meses + n_Dias) <> 0 Then
              a_Valores = Array(ps_ClsPlanilla, s_PeriodoCts, txtCodigo.Text, Trim(porstRecordset!codpsn), Trim(cmbejercicio), Trim(Left(cmbPeriodo, 2)), CInt(n_Annos), CInt(n_Meses), CInt(n_Dias), Format(s_FechaIni, s_FmtFechMysql_0), Format(porstRecordset!fechafin, s_FmtFechMysql_0), Format(porstRecordset!fechaven, s_FmtFechMysql_0), Format(porstRecordset!fechacan, s_FmtFechMysql_0), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
              If Not Records_Ins("plctsmovimiento", a_Campos, a_Valores, a_Tipos) Then GoTo Error
            End If
          End If
          porstRecordset.MoveNext
        Wend
        porstRecordset.Close
      End If
    End If
  Else
    If Not Records_Upd("plctsperiodosub", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fCtsPeriodoSub.dcaRegistro, fCtsPeriodoSub.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fCtsPeriodoSub.dcaRegistro.Recordset.Find ("subcts='" & s_Registro & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtCodigo.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  Set porstBusqueda = Nothing
  
End Sub
Private Sub dtpFechas_LostFocus(Index As Integer)
  Dim n_Dias  As Long
  Dim n_Annos As Integer, n_Meses As Integer
  
  If Index = 0 Then Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha final debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  n_Dias = gdl_Funcion.NumeroDias360(dtpFechas(1), dtpFechas(0), dtpFechas(1))
  ' Obtengo el numero de años
  If n_Dias >= 360 Then
    n_Annos = CInt(n_Dias \ 360)
    n_Dias = (n_Dias - (360 * n_Annos))
  End If
  ' Obtengo el numero de meses
  If n_Dias < 360 Then
    n_Meses = CInt(n_Dias \ 30)
    n_Dias = (n_Dias - (n_Meses * 30))
  End If
  txtnumero(0).Text = CInt(n_Annos)
  txtnumero(1).Text = CInt(n_Meses)
  txtnumero(2).Text = CInt(n_Dias)

End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 5480: Me.Width = 7830
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Sub Periodos de Provisión"
  lblTitle = "Sub Periodo de Provisión"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fCtsPeriodoSub.Tag

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 3)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&n", "&e", "&m")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento

  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento

  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True

  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0

  ' Verifico si existen Registros
  l_ExistRecord = (fCtsPeriodoSub.dcaRegistro.Recordset.EOF Or fCtsPeriodoSub.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fCtsPeriodoSub.dcaRegistro.Recordset!subcts
  
  ' Configuro los listados, datos adicionales
  cmbPeriodo.Clear
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  cmbejercicio.Clear
  For n_Index = 1 To 5: cmbejercicio.AddItem Choose(n_Index, Trim(ps_Anyo - 2), Trim(ps_Anyo - 1), ps_Anyo, Trim(ps_Anyo + 1), Trim(ps_Anyo + 2)): Next n_Index

  ' Carga los datos en el formulario
  ShowScreen

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub mskFechaCan_GotFocus()
  gdl_Procedure.MarcaGet mskFechaCan
End Sub
Private Sub mskFechaCan_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFechaCan_Validate(Cancel As Boolean)
  If mskFechaCan.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFechaCan, 1900
  End If
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub txtCodigo_GotFocus()
  gdl_Procedure.MarcaGet txtCodigo
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If txtCodigo = "" Then
      Beep
      MsgBox "Debe Ingresar el Código del " & lblTitle, vbExclamation
      txtCodigo.SetFocus
    Else
      txtDescripcion.SetFocus
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txtDescripcion_GotFocus()
  gdl_Procedure.MarcaGet txtDescripcion
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtnumero_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtnumero(Index)
End Sub
Private Sub txtnumero_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNumero_Validate(Index As Integer, Cancel As Boolean)
  txtnumero(Index).Text = IIf(Not IsNumeric(txtnumero(Index).Text), 0, txtnumero(Index).Text)
  txtnumero(Index).Text = IIf(CInt(txtnumero(Index).Text) < 0, 0, txtnumero(Index).Text)
  txtnumero(Index).Text = CInt(txtnumero(Index).Text)
End Sub
