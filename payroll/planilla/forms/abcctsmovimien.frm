VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcCtsMovimiento 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4875
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "abcctsmovimien.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   7740
   Begin TabDlg.SSTab tabRegister 
      Height          =   3720
      Left            =   75
      TabIndex        =   42
      Top             =   600
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6562
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
      TabPicture(0)   =   "abcctsmovimien.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "shpCuadro(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "shpCuadro(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDato(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDato(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblHelp(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblHelp(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdHelp(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdHelp(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "frmCuadro(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frmCuadro(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNumero(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNumero(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNumero(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtMeses"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtEjercicio"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtPeriodo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAnos"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      Begin VB.TextBox txtAnos 
         Height          =   280
         Left            =   285
         TabIndex        =   7
         Top             =   1360
         Width           =   780
      End
      Begin VB.TextBox txtPeriodo 
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   1395
         TabIndex        =   4
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox txtEjercicio 
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   1395
         TabIndex        =   1
         Top             =   195
         Width           =   900
      End
      Begin VB.TextBox txtMeses 
         Height          =   280
         Left            =   1260
         TabIndex        =   8
         Top             =   1360
         Width           =   1590
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   1
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1360
         Width           =   800
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   0
         Left            =   3315
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1360
         Width           =   800
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   2
         Left            =   5535
         MaxLength       =   4
         TabIndex        =   14
         Top             =   1360
         Width           =   800
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1530
         Index           =   1
         Left            =   4650
         TabIndex        =   24
         Top             =   1800
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   2699
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
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   25
            Top             =   405
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "P&endiente"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   705
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Provisionado"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   27
            Top             =   1005
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Cancelado"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1530
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1800
         Width           =   3945
         _Version        =   65536
         _ExtentX        =   6959
         _ExtentY        =   2699
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
            Left            =   270
            TabIndex        =   17
            Top             =   450
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            Format          =   138215425
            CurrentDate     =   37515
         End
         Begin MSComCtl2.DTPicker dtpFechas 
            DataField       =   "1530"
            Height          =   300
            Index           =   1
            Left            =   2340
            TabIndex        =   19
            Top             =   450
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   140705793
            CurrentDate     =   37515
         End
         Begin MSMask.MaskEdBox mskFechaCan 
            DataField       =   "1530"
            Height          =   300
            Left            =   2460
            TabIndex        =   23
            Top             =   1095
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
            Left            =   390
            TabIndex        =   21
            Top             =   1095
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   140705793
            CurrentDate     =   37515
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelación :"
            DataField       =   "1530"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   2460
            TabIndex        =   22
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   390
            TabIndex        =   20
            Top             =   840
            Width           =   1005
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   660
            Index           =   2
            Left            =   135
            Shape           =   4  'Rounded Rectangle
            Top             =   810
            Width           =   3675
         End
         Begin VB.Label lblDato 
            Caption         =   "Final :"
            DataField       =   "1530"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   2340
            TabIndex        =   18
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Inicio :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   270
            TabIndex        =   16
            Top             =   210
            Width           =   1005
         End
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   2370
         TabIndex        =   43
         Top             =   195
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   1
         Left            =   2370
         TabIndex        =   44
         Top             =   540
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Height          =   180
         Index           =   1
         Left            =   2730
         TabIndex        =   5
         Top             =   585
         Width           =   180
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Height          =   180
         Index           =   0
         Left            =   2730
         TabIndex        =   2
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sub-Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   3
         Top             =   585
         Width           =   1005
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   1
         Left            =   195
         Shape           =   4  'Rounded Rectangle
         Top             =   135
         Width           =   6405
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Meses :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4440
         TabIndex        =   11
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Años :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3315
         TabIndex        =   9
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dias :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5535
         TabIndex        =   13
         Top             =   1065
         Width           =   795
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   750
         Index           =   0
         Left            =   3060
         Shape           =   4  'Rounded Rectangle
         Top             =   1005
         Width           =   3540
      End
      Begin VB.Label lblDato 
         Caption         =   "Ejercicio / Mes Remuneración :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   6
         Top             =   1065
         Width           =   2565
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
         Picture         =   "abcctsmovimien.frx":0028
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
         Picture         =   "abcctsmovimien.frx":0044
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
      Top             =   4365
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
         Picture         =   "abcctsmovimien.frx":0060
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
         Picture         =   "abcctsmovimien.frx":007C
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
         Picture         =   "abcctsmovimien.frx":0098
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
         Picture         =   "abcctsmovimien.frx":00B4
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   3720
      Index           =   0
      Left            =   6960
      TabIndex        =   37
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6562
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
         Top             =   900
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsmovimien.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   40
         Tag             =   "0"
         Top             =   1710
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsmovimien.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   41
         Tag             =   "0"
         Top             =   2490
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   12632256
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcctsmovimien.frx":0108
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2595
      TabIndex        =   45
      Top             =   495
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   4233
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2117"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "fAbcCtsMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private s_Registro As String                           ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins)
  cmdHelp(1).Enabled = (Me.Tag = s_MdoData_Ins)

End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtEjercicio, "", Me.Tag, False, fCtsMovimiento.dcaRegistro.Recordset!pdocts.DefinedSize
    gdl_Procedure.EditText "PK", txtPeriodo, "", Me.Tag, False, fCtsMovimiento.dcaRegistro.Recordset!subcts.DefinedSize
    gdl_Procedure.EditText "PK", txtAnos, "", s_MdoData_Upd, True, fCtsMovimiento.dcaRegistro.Recordset!pdoano.DefinedSize
    gdl_Procedure.EditText "PK", txtMeses, "", s_MdoData_Upd, True, 15
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
  Else
    gdl_Procedure.EditText "PK", txtEjercicio, fCtsMovimiento.dcaRegistro.Recordset!pdocts, Me.Tag, True, fCtsMovimiento.dcaRegistro.Recordset!pdocts.DefinedSize
    gdl_Procedure.EditText "PK", txtPeriodo, fCtsMovimiento.dcaRegistro.Recordset!subcts, Me.Tag, True, fCtsMovimiento.dcaRegistro.Recordset!subcts.DefinedSize
    gdl_Procedure.EditText "PK", txtAnos, fCtsMovimiento.dcaRegistro.Recordset!pdoano, Me.Tag, True, fCtsMovimiento.dcaRegistro.Recordset!pdoano
    n_Index = (fCtsMovimiento.dcaRegistro.Recordset!pdomes)
    txtMeses.Text = Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre")
    gdl_Procedure.EditText "PK", txtMeses, txtMeses.Text, Me.Tag, True, 15
    gdl_Procedure.EditText "AT", txtnumero(0), CInt(fCtsMovimiento.dcaRegistro.Recordset!numeroanos), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtnumero(1), CInt(fCtsMovimiento.dcaRegistro.Recordset!numeromeses), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditText "AT", txtnumero(2), CInt(fCtsMovimiento.dcaRegistro.Recordset!numerodias), Me.Tag, False, 2, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), fCtsMovimiento.dcaRegistro.Recordset!fechaini, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), fCtsMovimiento.dcaRegistro.Recordset!fechafin, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(2), fCtsMovimiento.dcaRegistro.Recordset!fechaven, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditMask "AT", mskFechaCan, IIf(IsNull(fCtsMovimiento.dcaRegistro.Recordset!fechacan), "", fCtsMovimiento.dcaRegistro.Recordset!fechacan), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fCtsMovimiento.dcaRegistro.Recordset!estadomov = s_Estado_Ina), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fCtsMovimiento.dcaRegistro.Recordset!estadomov = s_Estado_Act), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(2), (fCtsMovimiento.dcaRegistro.Recordset!estadomov = s_Estado_Blq), Me.Tag, False
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtEjercicio, "EC")
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, Trim(txtEjercicio.Text) & "|" & Trim(txtPeriodo.Text), "SC")

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_PeriodoCts As String, s_Personal As String
  
  ' Valido que el peiodo no se eencuentre procesado
  If (optEstado(1).Value Or optEstado(2).Value) And Index <> 0 Then Beep: MsgBox "Movimiento No se puede Actualizar se encuentra Provisionado - Cancelado", vbExclamation: Me.Tag = s_MdoData_Vis: Exit Sub
  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtEjercicio.SetFocus
  ElseIf Index = 2 Then
   txtnumero(0).SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim$(lblHelp(0)) & "'/'" & Trim$(lblHelp(1)) & "'  ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a actualizar
    s_Registro = txtEjercicio.Text
    s_PeriodoCts = txtPeriodo.Text
    s_Personal = Trim(fCtsMovimiento.dcaRegistro.Recordset!codpsn)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "pdocts", "subcts", "codpsn")
    a_Valores = Array(ps_ClsPlanilla, s_Registro, s_PeriodoCts, s_Personal)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plctsmovimiento", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fCtsMovimiento.dcaRegistro, fCtsMovimiento.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fCtsMovimiento.dcaRegistro.Recordset.EOF And fCtsMovimiento.dcaRegistro.Recordset.BOF) Or fCtsMovimiento.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fCtsMovimiento.dcaRegistro.Recordset.Find ("cPrimaryKey >= '" & s_Registro & s_PeriodoCts & "'")
      If fCtsMovimiento.dcaRegistro.Recordset.EOF Then fCtsMovimiento.dcaRegistro.Recordset.MoveLast
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
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0      ' Periodos de cts
    tdbHelp.Columns(0).DataField = "pdocts": tdbHelp.Columns(1).DataField = "descricts"
    tdbHelp.Caption = "Periodo de CTS"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ced", "pdocts", s_Estado_Ina & ps_ClsPlanilla, "")
   Case 1       ' Sub periodo de cts
    tdbHelp.Columns(0).DataField = "subcts": tdbHelp.Columns(1).DataField = "descrisub"
    tdbHelp.Caption = "Sub periodo CTS"
    s_Registro = ps_ClsPlanilla & txtEjercicio.Text
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("sed", "subcts", s_Estado_Blq & s_Registro, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + (Choose(Index + 1, cmdHelp(Index).Top, 750, cmdHelp(Index).Top, 750, 850) + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero inicial, anterior, siguiente o final
  Select Case Index
   Case 0: fCtsMovimiento.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fCtsMovimiento.dcaRegistro.Recordset.BOF Then fCtsMovimiento.dcaRegistro.Recordset.MovePrevious
           If fCtsMovimiento.dcaRegistro.Recordset.BOF Then fCtsMovimiento.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fCtsMovimiento.dcaRegistro.Recordset.EOF Then fCtsMovimiento.dcaRegistro.Recordset.MoveNext
           If fCtsMovimiento.dcaRegistro.Recordset.EOF Then fCtsMovimiento.dcaRegistro.Recordset.MoveLast
   Case 3: fCtsMovimiento.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_PeriodoCts As String, s_Personal As String, s_FechaIni As String
  Dim s_Estado As String * 1, s_DesAusenciaBF As String
  Dim n_DiasCTS As Long, n_DiasIngreso As Long, n_DiasAusencia As Long
  Dim d_FechaIngreso As Date
  
  ' Realizo las validaciones de los campos a actualizar
  If txtEjercicio = "" Then Beep: MsgBox "Debe Ingresar el Periodo " & lblTitle, vbExclamation: txtEjercicio.SetFocus: Exit Sub
  If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Sub Periodo " & lblTitle, vbExclamation: txtPeriodo.SetFocus: Exit Sub
  If Not IsNumeric(txtnumero(0).Text) Or CInt(txtnumero(0).Text) < 0 Then Beep: MsgBox "Numero de años  no valido; verifique", vbExclamation: txtnumero(0).SetFocus: Exit Sub
  If Not IsNumeric(txtnumero(1).Text) Or CInt(txtnumero(1).Text) < 0 Or CInt(txtnumero(1).Text) > 11 Then Beep: MsgBox "Numero de meses no valido; verifique", vbExclamation: txtnumero(1).SetFocus: Exit Sub
  If Not IsNumeric(txtnumero(2).Text) Or CInt(txtnumero(2).Text) < 0 Or CInt(txtnumero(1).Text) > 29 Then Beep: MsgBox "Numero de dias no valido; verifique", vbExclamation: txtnumero(2).SetFocus: Exit Sub
  If txtAnos = "" Then Beep: MsgBox "Selecciono el Ejercicio de Remuneración", vbExclamation: txtPeriodo.SetFocus: Exit Sub
  If txtMeses = "" Then Beep: MsgBox "Selecciono el mes de Remuneración", vbExclamation: txtPeriodo.SetFocus: Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha final debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  If Not (dtpFechas(2) >= dtpFechas(1)) Then Beep: MsgBox "Fecha vencimiento debe ser mayor o igual que la fecha Final", vbExclamation: dtpFechas(2).SetFocus: Exit Sub
  If (optEstado(0).Value Or optEstado(1).Value) And mskFechaCan.ClipText <> "" Then Beep: MsgBox "No debe ngresar la Fecha de Cancelación", vbExclamation: mskFechaCan.SetFocus: Exit Sub
  If optEstado(2).Value And mskFechaCan.ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha de Cancelación", vbExclamation: mskFechaCan.SetFocus: Exit Sub
  If mskFechaCan.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFechaCan, 1900) Then mskFechaCan.SetFocus: Exit Sub
  End If
  If Not (Trim(txtAnos.Text) & Left(txtMeses, 2) >= Format(dtpFechas(0), "yyyymm") And Trim(txtAnos.Text) & Left(txtMeses, 2) <= Format(dtpFechas(1), "yyyymm")) Then Beep: MsgBox "Mes debe ser dentro del rango de las fechas", vbExclamation: txtPeriodo.SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, s_Estado_Ina, IIf(optEstado(1).Value, s_Estado_Act, s_Estado_Blq))
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = txtEjercicio.Text
  s_PeriodoCts = txtPeriodo.Text
  s_Personal = Trim(o_PvsComxTieSer.dcaRegistro.Recordset!codpsn)
  d_FechaIngreso = Format(o_PvsComxTieSer.dcaRegistro.Recordset!fecingreso, s_FormatoFecha)
  
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "pdocts", "subcts", "codpsn", "pdoano", "pdomes", "numeroanos", "numeromeses", "numerodias", "fechaini", "fechafin", "fechaven", "fechacan", "estadomov", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "pdocts", "subcts", "codpsn")
  
  ' Obtengo la fecha inicial del periodo
  s_FechaIni = Format(dtpFechas(0).Value, s_FormatoFecha)
  s_FechaIni = Format(IIf(Format(s_FechaIni, "yyyymmdd") >= Format(d_FechaIngreso, "yyyymmdd"), s_FechaIni, d_FechaIngreso), s_FormatoFecha)
  s_FechaIni = Format(IIf(Format(s_FechaIni, "yyyymmdd") > Format(d_FechaIngreso, "yyyymmdd"), DateAdd("m", 1, "01" & Mid(s_FechaIni, 3)), s_FechaIni), s_FormatoFecha)
  dtpFechas(0).Value = s_FechaIni
  
  n_DiasCTS = 0
  ' Obtengo los años, meses y dias
  n_DiasCTS = gdl_Funcion.NumeroDias360Nuevo(Format(dtpFechas(1).Value, s_FormatoFecha), s_FechaIni, Format(dtpFechas(1).Value, s_FormatoFecha))
  
  ' Vaido si trabajador laboral por lo menos un mes
  n_DiasIngreso = DateDiff("d", d_FechaIngreso, dtpFechas(1).Value) + 1
  If Not ((n_DiasIngreso >= 30 Or (Format(dtpFechas(1).Value, "yyyymmdd") >= Format(DateAdd("m", 1, d_FechaIngreso), "yyyymmdd"))) And n_DiasCTS >= 30) Then Beep: MsgBox "Rango de periodo no valido para el trabajador; Verfique", vbExclamation: dtpFechas(1).SetFocus: GoTo Finalizar
  s_DesAusenciaBF = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, "", ps_Anyo, "RA")
  ' Resta Ausencias encontradas en el período (Si el parámetro así lo indica)
  n_DiasAusencia = 0
  If s_DesAusenciaBF = s_Estado_Act Then
    n_DiasAusencia = gdl_Funcion.DiasAusenciaBS(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, s_Personal, s_FechaIni, Format(dtpFechas(1).Value, s_FormatoFecha))
  End If
  n_DiasCTS = n_DiasCTS - n_DiasAusencia
  
  txtnumero(2).Text = 0
  txtnumero(1).Text = 0
  ' Obtengo el numero de años
  If n_DiasCTS >= 360 Then
    txtnumero(2).Text = CLng(n_DiasCTS \ 360)
    n_DiasCTS = (n_DiasCTS - (360 * CLng(txtnumero(2).Text)))
  End If
  ' Obtengo el numero de meses
  If n_DiasCTS < 360 Then
    txtnumero(1).Text = CLng(n_DiasCTS \ 30)
    n_DiasCTS = (n_DiasCTS - (CLng(txtnumero(1).Text) * 30))
  End If
  
  ' Valido y realizo la actualizacion de movimientos de cts
  If (CLng(txtnumero(0).Text) + CLng(txtnumero(1).Text) + CLng(txtnumero(2).Text)) <= 0 Then Beep: MsgBox "Dias, meses y años valido; verifique", vbExclamation: txtnumero(0).SetFocus: GoTo Finalizar
  a_Valores = Array(ps_ClsPlanilla, Trim(txtEjercicio.Text), Trim(txtPeriodo.Text), s_Personal, Trim(txtAnos.Text), Trim(Left(txtMeses.Text, 2)), CInt(txtnumero(0).Text), CInt(txtnumero(1).Text), CInt(txtnumero(2).Text), Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), Format(dtpFechas(2), s_FmtFechMysql_0), Format(mskFechaCan, s_FmtFechMysql_0), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plctsmovimiento", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plctsmovimiento", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fCtsMovimiento.dcaRegistro, fCtsMovimiento.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fCtsMovimiento.dcaRegistro.Recordset.Find ("cPrimaryKey='" & s_Registro & s_PeriodoCts & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtEjercicio.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  
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
  Me.Height = 5360: Me.Width = 7830
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Movimientos de Provisión"
  lblTitle = "Movimiento de Provisión"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fCtsMovimiento.Tag

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
  l_ExistRecord = (fCtsMovimiento.dcaRegistro.Recordset.EOF Or fCtsMovimiento.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fCtsMovimiento.dcaRegistro.Recordset!cPrimaryKey
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcts", "descricts")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 734.7402, 3465.071)
      aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify)
      aElemento(n_Index, 4) = Choose(n_Index + 1, "", "")
      aElemento(n_Index, 5) = Choose(n_Index + 1, False, False)
      aElemento(n_Index, 6) = Choose(n_Index + 1, True, True)
      aElemento(n_Index, 7) = Choose(n_Index + 1, "", "")
      aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop)
      aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0)
  Next n_Index
  
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
      aElementos(n_Index, 0) = ""
      aElementos(n_Index, 1) = n_BackColorHelp#: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbHelp, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Cálculo", 2
  ']

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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0         ' Periodo de cts
    txtEjercicio = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtEjercicio.SetFocus
   Case 1         ' Sub periodo de cts
    txtPeriodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Periodo de cts
    s_Sql = gdl_Funcion.HelpTablas("pcs", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
   Case 1     ' Sub periodo de cts
    s_Registro = ps_ClsPlanilla & Trim(txtEjercicio.Text)
    s_Sql = gdl_Funcion.HelpTablas("sxe", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then porstHelp.Requery
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim porstClone As ADODB.Recordset
  Dim n_Columna As Integer, s_Criterio As String

  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    ' Conformo la cadena de ayuda
    s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    Set porstClone = porstHelp.Clone()
    
    n_Columna = tdbHelp.Col
    s_Criterio = tdbHelp.Columns(n_Columna).DataField & " >= '" & s_SqlHelp & "'"
    porstClone.Find s_Criterio, 0, adSearchForward, 0
    If Not (porstClone.BOF Or porstClone.EOF) Then
      porstHelp.Bookmark = porstClone.Bookmark
    End If
    porstClone.Close
    Set porstClone = Nothing
  Else
      s_SqlHelp = ""
  End If

End Sub
Private Sub tdbHelp_LostFocus()
  tdbHelp.Visible = False
End Sub

Private Sub txtAnos_GotFocus()
  gdl_Procedure.MarcaGet txtMeses
End Sub
Private Sub txtAnos_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtEjercicio_GotFocus()
  gdl_Procedure.MarcaGet txtEjercicio
End Sub
Private Sub txtEjercicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtEjercicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtEjercicio_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtEjercicio, "EC")
End Sub
Private Sub txtEjercicio_Validate(Cancel As Boolean)
  
  If Me.Tag = s_MdoData_Ins Then
    lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, Trim(txtEjercicio.Text) & "|" & Trim(txtPeriodo.Text), "SC")
    If Not (lblHelp(1) = "???" Or lblHelp(1) = "") Then Exit Sub
    ' Inicializo controles
    txtPeriodo.Text = ""
    lblHelp(1) = ""
    txtAnos.Text = ""
    txtMeses.Text = ""
    txtnumero(0).Text = CInt(0)
    txtnumero(1).Text = CInt(0)
    txtnumero(2).Text = CInt(0)
    mskFechaCan.ToolTipText = Format("", s_FormatoFecha)
    optEstado(0).Value = True
    optEstado(1).Value = False
  End If

End Sub
Private Sub txtMeses_GotFocus()
  gdl_Procedure.MarcaGet txtMeses
End Sub
Private Sub txtMeses_KeyPress(KeyAscii As Integer)
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
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, Trim(txtEjercicio.Text) & "|" & Trim(txtPeriodo.Text), "SC")
End Sub
Private Sub txtPeriodo_Validate(Cancel As Boolean)

  If Trim(txtPeriodo.Text) <> "" And Me.Tag = s_MdoData_Ins Then
    ' Inicializo controles
    lblHelp(1) = ""
    txtAnos.Text = ""
    txtMeses.Text = ""
    txtnumero(0).Text = CInt(0)
    txtnumero(1).Text = CInt(0)
    txtnumero(2).Text = CInt(0)
    ' Cargo registros del sub periodo
    s_Sql = "SELECT descrisub, pdoano, pdomes, numeroanos, numeromeses, numerodias, "
    s_Sql = s_Sql & "fechaini, fechafin, fechacan, estadosub "
    s_Sql = s_Sql & "FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts= '" & Trim(txtEjercicio.Text) & "' "
    s_Sql = s_Sql & "AND subcts= '" & Trim(txtPeriodo.Text) & "' "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      ' Inicializo controles
      lblHelp(1) = Trim(porstRecordset!descrisub)
      txtAnos.Text = Trim(porstRecordset!pdoano)
      n_Index = CInt(porstRecordset!pdomes)
      txtMeses.Text = Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre")
      txtnumero(0).Text = CInt(porstRecordset!numeroanos)
      txtnumero(1).Text = CInt(porstRecordset!numeromeses)
      txtnumero(2).Text = CInt(porstRecordset!numerodias)
      dtpFechas(0) = Format(porstRecordset!fechaini, s_FormatoFecha)
      dtpFechas(1) = Format(porstRecordset!fechafin, s_FormatoFecha)
      mskFechaCan.ToolTipText = Format(porstRecordset!fechacan, s_FormatoFecha)
      s_Registro = Trim(porstRecordset!estadosub)
      optEstado(0).Value = (s_Registro = s_Estado_Ina)
      optEstado(1).Value = (s_Registro = s_Estado_Act)
    End If
  End If

End Sub
