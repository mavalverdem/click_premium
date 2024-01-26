VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fPrmPlanillaAfp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5610
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   8415
   Icon            =   "prmplanillafp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8415
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
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
         Left            =   7560
         TabIndex        =   48
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
         Picture         =   "prmplanillafp.frx":000C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   7170
         TabIndex        =   49
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmplanillafp.frx":0028
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
         Left            =   510
         TabIndex        =   44
         Top             =   120
         Width           =   6000
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   45
      Top             =   5100
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
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
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   4410
      Left            =   75
      TabIndex        =   46
      Top             =   600
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   7779
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Planilla"
      TabPicture(0)   =   "prmplanillafp.frx":0044
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDetalle(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDetalle(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmCuadro(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Parametros"
      TabPicture(1)   =   "prmplanillafp.frx":0060
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblHelp(0)"
      Tab(1).Control(1)=   "lblDato(0)"
      Tab(1).Control(2)=   "frmCuadro(6)"
      Tab(1).Control(3)=   "frmCuadro(1)"
      Tab(1).Control(4)=   "frmCuadro(0)"
      Tab(1).Control(5)=   "cmdHelp(0)"
      Tab(1).Control(6)=   "txtRemunera"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtRemunera 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -74745
         TabIndex        =   28
         Top             =   510
         Width           =   980
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   -73695
         TabIndex        =   50
         Top             =   510
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2970
         Index           =   0
         Left            =   -74880
         TabIndex        =   29
         Top             =   945
         Width           =   4000
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   5239
         _StockProps     =   14
         Caption         =   " Fondo de Pensiones "
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
         Begin VB.TextBox txtAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   3
            Left            =   180
            TabIndex        =   37
            Top             =   2520
            Width           =   980
         End
         Begin VB.TextBox txtAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   2
            Left            =   180
            TabIndex        =   35
            Top             =   1875
            Width           =   980
         End
         Begin VB.TextBox txtAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   1
            Left            =   180
            TabIndex        =   33
            Top             =   1215
            Width           =   980
         End
         Begin VB.TextBox txtAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   0
            Left            =   180
            TabIndex        =   31
            Top             =   570
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   1
            Left            =   1220
            TabIndex        =   52
            Top             =   570
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   2
            Left            =   1220
            TabIndex        =   54
            Top             =   1215
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   3
            Left            =   1215
            TabIndex        =   56
            Top             =   1875
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   4
            Left            =   1215
            TabIndex        =   58
            Top             =   2520
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "Aporte Empleador :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   36
            Top             =   2235
            Width           =   2730
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
            Height          =   195
            Index           =   4
            Left            =   1575
            TabIndex        =   59
            Top             =   2565
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Aporte Voluntario sin Fin Previsional :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   34
            Top             =   1590
            Width           =   2730
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
            Height          =   195
            Index           =   3
            Left            =   1575
            TabIndex        =   57
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Aporte Voluntario con Fin Previsional :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   32
            Top             =   930
            Width           =   2730
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
            Height          =   195
            Index           =   2
            Left            =   1580
            TabIndex        =   55
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Aporte Obligatorio :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   30
            Top             =   285
            Width           =   2730
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
            Height          =   195
            Index           =   1
            Left            =   1580
            TabIndex        =   53
            Top             =   615
            Width           =   195
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1890
         Index           =   1
         Left            =   -70845
         TabIndex        =   38
         Top             =   960
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   3334
         _StockProps     =   14
         Caption         =   " Retenciones y Retribuciones "
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
         Begin VB.TextBox txtRetencion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   0
            Left            =   180
            TabIndex        =   40
            Top             =   570
            Width           =   980
         End
         Begin VB.TextBox txtRetencion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   1
            Left            =   180
            TabIndex        =   42
            Top             =   1215
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   5
            Left            =   1220
            TabIndex        =   60
            Top             =   570
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   6
            Left            =   1220
            TabIndex        =   61
            Top             =   1215
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
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
            Height          =   195
            Index           =   5
            Left            =   1580
            TabIndex        =   63
            Top             =   615
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Seguros :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   39
            Top             =   285
            Width           =   2730
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
            Height          =   195
            Index           =   6
            Left            =   1580
            TabIndex        =   62
            Top             =   1260
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Comisión % sobre R.A. :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   41
            Top             =   930
            Width           =   2730
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2820
         Index           =   2
         Left            =   110
         TabIndex        =   2
         Top             =   1125
         Width           =   8010
         _Version        =   65536
         _ExtentX        =   14129
         _ExtentY        =   4974
         _StockProps     =   14
         Caption         =   " Datos Cabecera "
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
         Begin VB.TextBox txtNroHoja 
            Height          =   280
            Left            =   195
            TabIndex        =   4
            Top             =   540
            Width           =   1155
         End
         Begin VB.TextBox txtPeriodo 
            Height          =   280
            Left            =   1770
            TabIndex        =   6
            Top             =   555
            Width           =   885
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   285
            Left            =   3705
            TabIndex        =   8
            Top             =   555
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Format          =   141688833
            CurrentDate     =   37515
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   1500
            Index           =   4
            Left            =   135
            TabIndex        =   13
            Top             =   1100
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   2646
            _StockProps     =   14
            Caption         =   " Fondo de Pensiones  "
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
            Begin VB.TextBox txtBanco 
               Height          =   300
               Index           =   0
               Left            =   1290
               TabIndex        =   19
               Top             =   1050
               Width           =   495
            End
            Begin VB.TextBox txtInteres 
               Height          =   300
               Index           =   0
               Left            =   1290
               TabIndex        =   15
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox txtCheque 
               Height          =   300
               Index           =   0
               Left            =   1290
               TabIndex        =   17
               Top             =   705
               Width           =   1200
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   7
               Left            =   1845
               TabIndex        =   64
               Top             =   1050
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Banco :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   180
               TabIndex        =   18
               Top             =   1095
               Width           =   1005
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Intereses :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   12
               Left            =   180
               TabIndex        =   14
               Top             =   405
               Width           =   1005
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
               Height          =   195
               Index           =   7
               Left            =   2205
               TabIndex        =   65
               Top             =   1095
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Cheque Nro. :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   11
               Left            =   180
               TabIndex        =   16
               Top             =   750
               Width           =   1005
            End
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   1500
            Index           =   5
            Left            =   4050
            TabIndex        =   20
            Top             =   1100
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   2646
            _StockProps     =   14
            Caption         =   "Retenciones y Retribuciones "
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
            Begin VB.TextBox txtCheque 
               Height          =   300
               Index           =   1
               Left            =   1290
               TabIndex        =   24
               Top             =   705
               Width           =   1200
            End
            Begin VB.TextBox txtInteres 
               Height          =   300
               Index           =   1
               Left            =   1290
               TabIndex        =   22
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox txtBanco 
               Height          =   300
               Index           =   1
               Left            =   1290
               TabIndex        =   26
               Top             =   1050
               Width           =   495
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   8
               Left            =   1845
               TabIndex        =   66
               Top             =   1050
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Cheque Nro. :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   13
               Left            =   180
               TabIndex        =   23
               Top             =   750
               Width           =   1005
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
               Height          =   195
               Index           =   8
               Left            =   2205
               TabIndex        =   67
               Top             =   1095
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Intereses :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   10
               Left            =   180
               TabIndex        =   21
               Top             =   405
               Width           =   1005
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Banco :"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   180
               TabIndex        =   25
               Top             =   1095
               Width           =   1005
            End
         End
         Begin Threed.SSFrame frmCuadro 
            Height          =   555
            Index           =   3
            Left            =   5130
            TabIndex        =   9
            Top             =   255
            Width           =   2760
            _Version        =   65536
            _ExtentX        =   4868
            _ExtentY        =   979
            _StockProps     =   14
            Caption         =   " Forma de Pago  "
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
            Begin Threed.SSOption optFormaPago 
               Height          =   180
               Index           =   0
               Left            =   225
               TabIndex        =   10
               Top             =   270
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   317
               _StockProps     =   78
               Caption         =   "&Efectivo"
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
            Begin Threed.SSOption optFormaPago 
               Height          =   180
               Index           =   1
               Left            =   1530
               TabIndex        =   11
               Top             =   270
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   317
               _StockProps     =   78
               Caption         =   "&Cheque"
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
         Begin Threed.SSCheck chkSinPago 
            Height          =   240
            Left            =   1770
            TabIndex        =   12
            Top             =   855
            Width           =   2250
            _Version        =   65536
            _ExtentX        =   3969
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Declaración sin pago  "
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
            Font3D          =   1
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Pago :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   16
            Left            =   3705
            TabIndex        =   7
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblDato 
            Caption         =   "Nro. Hoja :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   9
            Left            =   195
            TabIndex        =   3
            Top             =   285
            Width           =   780
         End
         Begin VB.Label lblDato 
            Caption         =   "Periodo que Devenga :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   15
            Left            =   1770
            TabIndex        =   5
            Top             =   300
            Width           =   1890
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1050
         Index           =   6
         Left            =   -70800
         TabIndex        =   68
         Top             =   2880
         Width           =   3930
         _Version        =   65536
         _ExtentX        =   6932
         _ExtentY        =   1852
         _StockProps     =   14
         Caption         =   "Ley 27252"
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
         Begin VB.TextBox txtAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   4
            Left            =   180
            TabIndex        =   69
            Top             =   570
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   9
            Left            =   1220
            TabIndex        =   70
            Top             =   570
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
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
            Height          =   195
            Index           =   9
            Left            =   1580
            TabIndex        =   72
            Top             =   615
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Aporte Obligatorio Empleador :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   14
            Left            =   180
            TabIndex        =   71
            Top             =   285
            Width           =   2730
         End
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
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
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
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
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDato 
         Caption         =   "Remuneración  Asegurable :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   -74745
         TabIndex        =   27
         Top             =   225
         Width           =   2070
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
         Height          =   195
         Index           =   0
         Left            =   -73335
         TabIndex        =   51
         Top             =   555
         Width           =   195
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   47
      Top             =   480
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
Attribute VB_Name = "fPrmPlanillaAfp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_ModoPll As String, s_ModoCfg As String        ' Modo de actualización de los datos
Private s_Mes As String                                 ' Mes de devengue
'[
Sub ShowScreen()
    
  ' Pestaña de Información de Planilla de Pensiones
  s_Sql = "SELECT pdoano, pdomes, codafp, nrohoja, sinpago, fechapago, formapago, interespension, "
  s_Sql = s_Sql & "chequepension, codbcopension, interesadmin, chequeadmin, codbcoadmin "
  s_Sql = s_Sql & "FROM plplanillafp "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND pdomes='" & s_Mes & "' "
  s_Sql = s_Sql & "AND codafp='" & s_Registro & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    s_ModoPll = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtNroHoja, gdl_Funcion.aTexto(porstRecordset!nrohoja), s_ModoPll, False, porstRecordset!nrohoja.DefinedSize
    gdl_Procedure.EditText "PK", txtPeriodo, s_Mes & "/" & ps_Anyo, s_MdoData_Vis, True, 7
    gdl_Procedure.EditDTPicker "AT", dtpFecha, porstRecordset!fechapago, s_ModoPll, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", optFormaPago(0), (porstRecordset!formapago = "E"), s_ModoPll, True
    gdl_Procedure.EditOptionCheck "AT", optFormaPago(1), (porstRecordset!formapago = "C"), s_ModoPll, True
    gdl_Procedure.EditOptionCheck "AT", chkSinPago, (porstRecordset!sinpago = s_Estado_Act), s_ModoPll, True
    gdl_Procedure.EditText "AT", txtInteres(0), FormatNumber(porstRecordset!interespension, 2), s_ModoPll, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtCheque(0), gdl_Funcion.aTexto(porstRecordset!chequepension), s_ModoPll, False, porstRecordset!chequepension.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco(0), gdl_Funcion.aTexto(porstRecordset!codbcopension), s_ModoPll, False, porstRecordset!codbcopension.DefinedSize
    gdl_Procedure.EditText "AT", txtInteres(1), FormatNumber(porstRecordset!interesadmin, 2), s_ModoPll, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtCheque(1), gdl_Funcion.aTexto(porstRecordset!chequeadmin), s_ModoPll, False, porstRecordset!chequeadmin.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco(1), gdl_Funcion.aTexto(porstRecordset!codbcoadmin), s_ModoPll, False, porstRecordset!codbcoadmin.DefinedSize
  Else
    s_ModoPll = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtNroHoja, "", s_ModoPll, False, porstRecordset!nrohoja.DefinedSize
    gdl_Procedure.EditText "PK", txtPeriodo, s_Mes & "/" & ps_Anyo, s_MdoData_Vis, True, 7
    gdl_Procedure.EditDTPicker "AT", dtpFecha, Date, s_ModoPll, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", optFormaPago(0), True, s_ModoPll, True
    gdl_Procedure.EditOptionCheck "AT", optFormaPago(1), False, s_ModoPll, True
    gdl_Procedure.EditOptionCheck "AT", chkSinPago, False, s_ModoPll, True
    gdl_Procedure.EditText "AT", txtInteres(0), FormatNumber(0, 2), s_ModoPll, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtCheque(0), "", s_ModoPll, False, porstRecordset!chequepension.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco(0), "", s_ModoPll, False, porstRecordset!codbcopension.DefinedSize
    gdl_Procedure.EditText "AT", txtInteres(1), FormatNumber(0, 2), s_ModoPll, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtCheque(1), "", s_ModoPll, False, porstRecordset!chequeadmin.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco(1), "", s_ModoPll, False, porstRecordset!codbcoadmin.DefinedSize
  End If
  
  lblHelp(7) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(0), "EB")
  lblHelp(8) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(1), "EB")

  ' Pestaña de Información de configuración
  s_Sql = "SELECT cpcremase, cpcapobli, cpcapovolsfp, cpcapovolcfp, cpc27252,cpcapoemp, cpcseguro, cpcporcen "
  s_Sql = s_Sql & "FROM plparametroafp "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    s_ModoCfg = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtRemunera, gdl_Funcion.aTexto(porstRecordset!cpcremase), s_ModoCfg, False, porstRecordset!cpcremase.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(0), gdl_Funcion.aTexto(porstRecordset!cpcapobli), s_ModoCfg, False, porstRecordset!cpcapobli.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(1), gdl_Funcion.aTexto(porstRecordset!cpcapovolsfp), s_ModoCfg, False, porstRecordset!cpcapovolsfp.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(2), gdl_Funcion.aTexto(porstRecordset!cpcapovolcfp), s_ModoCfg, False, porstRecordset!cpcapovolcfp.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(3), gdl_Funcion.aTexto(porstRecordset!cpcapoemp), s_ModoCfg, False, porstRecordset!cpcapoemp.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(4), gdl_Funcion.aTexto(porstRecordset!cpc27252), s_ModoCfg, False, porstRecordset!cpc27252.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(0), gdl_Funcion.aTexto(porstRecordset!cpcseguro), s_ModoCfg, False, porstRecordset!cpcseguro.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(1), gdl_Funcion.aTexto(porstRecordset!cpcporcen), s_ModoCfg, False, porstRecordset!cpcporcen.DefinedSize
  Else
    s_ModoCfg = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtRemunera, "", s_ModoCfg, False, porstRecordset!cpcremase.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(0), "", s_ModoCfg, False, porstRecordset!cpcapobli.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(1), "", s_ModoCfg, False, porstRecordset!cpcapovolsfp.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(2), "", s_ModoCfg, False, porstRecordset!cpcapovolcfp.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(3), "", s_ModoCfg, False, porstRecordset!cpcapoemp.DefinedSize
    gdl_Procedure.EditText "AT", txtAporte(4), "", s_ModoCfg, False, porstRecordset!cpc27252.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(0), "", s_ModoCfg, False, porstRecordset!cpcseguro.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(1), "", s_ModoCfg, False, porstRecordset!cpcporcen.DefinedSize
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera, "CP")
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(0), "CP")
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(1), "CP")
  lblHelp(3) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(2), "CP")
  lblHelp(4) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(3), "CP")
  lblHelp(9) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(4), "CP")
  lblHelp(5) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(0), "CP")
  lblHelp(6) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(1), "CP")

End Sub
']
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim nTop As Double, nLeft As Double
  
  nTop = (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  nLeft = (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Concepto de remuneración asegurable
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos de Ingresos"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Ina, "")
   Case 1, 2, 3, 5, 6, 9 ' Conceptos de descuento de pensiones
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Descuento"
    nTop = (frmCuadro(0).Top + IIf(Index = 3, 1300, cmdHelp(Index).Top) + (cmdHelp(Index).Height / 2))
    nLeft = (IIf(Index > 3, 2250, frmCuadro(0).Left) + cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Act, "")
   Case 4     ' Conceptos de planilla remuneracion
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos de Aportes"
    nTop = (frmCuadro(0).Top + 1450 + (cmdHelp(Index).Height / 2))
    nLeft = (frmCuadro(0).Left + cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Blq, "")
   Case 7, 8    ' Entidad de banco
    tdbHelp.Columns(0).DataField = "codbco": tdbHelp.Columns(1).DataField = "desbco"
    tdbHelp.Caption = "Entidad de Banco"
    nTop = (frmCuadro(2).Top + 1250 + (cmdHelp(Index).Height / 2))
    nLeft = (frmCuadro(2).Left + IIf(Index = 7, 2000, 3200) + (cmdHelp(Index).Width / 2))
    s_Sql = gdl_Funcion.HelpTablas("bco", "codbco", "", "")
End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + nTop
  tdbHelp.Left = tabRegister.Left + nLeft
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  Dim s_TipoPago As String * 1, s_SinPago As String * 1
  
  ' Primera pestaña (planilla)
  s_TipoPago = IIf(optFormaPago(0).Value, "E", "C")
  s_SinPago = IIf(chkSinPago.Value, s_Estado_Act, s_Estado_Ina)
  If txtNroHoja = "" Then Beep: MsgBox "Debe Ingresar el numero de hoja de planilla", vbExclamation: txtNroHoja.SetFocus: Exit Sub
  If Not (Format(dtpFecha, "yyyymm") >= (ps_Anyo & s_Mes)) Then Beep: MsgBox "Fecha de pago no valida; verificar", vbExclamation: dtpFecha.SetFocus: Exit Sub
  If CDec(txtInteres(0)) < 0 Then Beep: MsgBox "Valor de interes de Fondo de Pensión", vbExclamation: txtInteres(0).SetFocus: Exit Sub
  If txtCheque(0) = "" And s_TipoPago = "C" Then Beep: MsgBox "Debe Ingresar el Numero de Cheque del Fondo", vbExclamation: txtCheque(0).SetFocus: Exit Sub
  If txtBanco(0) = "" Then Beep: MsgBox "Debe Ingresar el banco del Fondo", vbExclamation: txtBanco(0).SetFocus: Exit Sub
  If lblHelp(7) = "???" Then Beep: MsgBox "Banco del Fondo no es valido; Verificar", vbExclamation: txtBanco(0).SetFocus: Exit Sub
  If CDec(txtInteres(1)) < 0 Then Beep: MsgBox "Valor de interes de Fondo de Retenciones", vbExclamation: txtInteres(1).SetFocus: Exit Sub
  If txtCheque(1) = "" And s_TipoPago = "C" Then Beep: MsgBox "Debe Ingresar el Numero de Cheque de Retenciones", vbExclamation: txtCheque(1).SetFocus: Exit Sub
  If txtBanco(1) = "" Then Beep: MsgBox "Debe Ingresar el banco de Retenciones", vbExclamation: txtBanco(1).SetFocus: Exit Sub
  If lblHelp(8) = "???" Then Beep: MsgBox "Banco de Retenciones no es valido; Verificar", vbExclamation: txtBanco(1).SetFocus: Exit Sub
  
  ' Segunda pestaña (configuración)
  If txtRemunera = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Asegurable", vbExclamation: txtRemunera.SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Concepto de Remuneración Asegurable no es valido; Verificar", vbExclamation: txtRemunera.SetFocus: Exit Sub
  If txtAporte(0) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Aporte Obligatorio", vbExclamation: txtAporte(0).SetFocus: Exit Sub
  If lblHelp(1) = "???" Then Beep: MsgBox "Concepto de Aporte Obligatorio no es valido; Verificar", vbExclamation: txtAporte(0).SetFocus: Exit Sub
  If lblHelp(2) = "???" Then Beep: MsgBox "Concepto de Aporte Volunt. S/F Previsional no es valido; Verificar", vbExclamation: txtAporte(1).SetFocus: Exit Sub
  If lblHelp(3) = "???" Then Beep: MsgBox "Concepto de Aporte Volunt. C/F Previsional no es valido; Verificar", vbExclamation: txtAporte(2).SetFocus: Exit Sub
  If lblHelp(4) = "???" Then Beep: MsgBox "Concepto de Aporte Empleador no es valido; Verificar", vbExclamation: txtAporte(3).SetFocus: Exit Sub
  If lblHelp(9) = "???" Then Beep: MsgBox "Concepto de Aporte Empleador Ley 27252 no es valido; Verificar", vbExclamation: txtAporte(4).SetFocus: Exit Sub
  For n_Index = 0 To 1
    If txtRetencion(n_Index) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de " & Choose(n_Index + 1, "Seguros", "Comisión % R.A."), vbExclamation: txtRetencion(n_Index).SetFocus: Exit Sub
    If lblHelp(n_Index + 5) = "???" Then Beep: MsgBox "Concepto de " & Choose(n_Index + 1, "Seguros", "Comisión % R.A.") & " no es valido; Verificar", vbExclamation: txtRetencion(n_Index).SetFocus: Exit Sub
  Next n_Index
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualización planilla
  a_Campos = Array("pdoano", "pdomes", "codafp", "nrohoja", "sinpago", "fechapago", "formapago", "interespension", "chequepension", "codbcopension", "interesadmin", "chequeadmin", "codbcoadmin", IIf(s_ModoPll = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(s_ModoPll = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, s_Mes, s_Registro, Trim(txtNroHoja.Text), s_SinPago, Format(dtpFecha, s_FmtFechMysql_0), s_TipoPago, CDec(txtInteres(0).Text), Trim(txtCheque(0).Text), Trim(txtBanco(0).Text), CDec(txtInteres(1).Text), Trim(txtCheque(1).Text), Trim(txtBanco(1).Text), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano", "pdomes", "codafp")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If s_ModoPll = s_MdoData_Ins Then
    If Not Records_Ins("plplanillafp", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plplanillafp", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  
  ' Creo los arreglos para la actualización parametros
  a_Campos = Array("pdoano", "cpcremase", "cpcapobli", "cpcapovolsfp", "cpcapovolcfp", "cpcapoemp", "cpc27252", "cpcseguro", "cpcporcen", IIf(s_ModoCfg = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(s_ModoCfg = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, txtRemunera.Text, txtAporte(0).Text, txtAporte(1).Text, txtAporte(2).Text, txtAporte(3).Text, txtAporte(4).Text, txtRetencion(0).Text, txtRetencion(1).Text, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano")
  ' Realizo el proceso de actualización de los registros
  If s_ModoCfg = s_MdoData_Ins Then
    If Not Records_Ins("plparametroafp", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plparametroafp", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Ubico el registro ingresado o actualizado
  ShowScreen
  txtRemunera.SetFocus
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 6090: Me.Width = 8500
  Me.Left = 3080: Me.Top = 2000
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Parametros de Planilla Afp"
  lblTitle = "Parametros"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Datos de Planilla
  s_Mes = Left(fReporPlanillAfp.cmbPeriodo, 2)
  s_Registro = Trim(fReporPlanillAfp.txtAfp)
  lblDetalle(0) = " Entidad de Pensión : " & Trim(fReporPlanillAfp.lblHelp(0)) & " "
  lblDetalle(1) = " Periodo : " & Trim(fReporPlanillAfp.cmbPeriodo) & "  " & ps_Anyo & " "
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "edit": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "aceptar"
  aElemento(0, 2) = "Actualizar Información de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdUpdate, aElemento
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcpc", "descpc")
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
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
End Sub
Private Sub optFormaPago_Click(Index As Integer, Value As Integer)
  If Index = 0 Then
    txtCheque(0) = "": txtCheque(1) = ""
  End If
  txtCheque(0).Locked = (Index = 0): txtCheque(1).Locked = (Index = 0)
End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0               ' Concepto de remuneración ganada
    txtRemunera = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtRemunera.SetFocus
   Case 1, 2, 3, 4  ' Concepto de aportes de fondo de pensiones
    txtAporte(n_IndexHelp - 1) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtAporte(n_IndexHelp - 1).SetFocus
   Case 5, 6            ' Concepto de retenciones
    txtRetencion(n_IndexHelp - 5) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtRetencion(n_IndexHelp - 5).SetFocus
   Case 7, 8            ' Entidad de banco
    txtBanco(n_IndexHelp - 7) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtBanco(n_IndexHelp - 7).SetFocus
   Case 9               ' Concepto de aportes de fondo de pensiones
    txtAporte(4) = tdbHelp.Columns(0).Value
    lblHelp(9) = tdbHelp.Columns(1).Value
    txtAporte(4).SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Concepto de remuneración asegurable
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina, "")
   Case 1, 2, 3, 5, 6, 9 ' Conceptos de descuento de pensiones
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Act, "")
   Case 4     ' Conceptos de planilla remuneracion
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Blq, "")
   Case 7, 8  ' Entidad de banco
    s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(ColIndex).DataField, "", "")
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
Private Sub txtAporte_GotFocus(Index As Integer)
   gdl_Procedure.MarcaGet txtAporte(Index)
End Sub
Private Sub txtAporte_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 1
End Sub
Private Sub txtAporte_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtAporte_LostFocus(Index As Integer)
   If Index = 4 Then
   lblHelp(9) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(Index), "CP")
   Else
   lblHelp(Index + 1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtAporte(Index), "CP")
   End If
End Sub
Private Sub txtBanco_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtBanco(Index)
End Sub
Private Sub txtBanco_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 7
End Sub
Private Sub txtBanco_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtBanco_LostFocus(Index As Integer)
  lblHelp(Index + 7) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco(Index), "EB")
End Sub
Private Sub txtCheque_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCheque(Index)
End Sub
Private Sub txtCheque_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtInteres_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtInteres(Index)
End Sub
Private Sub txtInteres_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtInteres_Validate(Index As Integer, Cancel As Boolean)
  txtInteres(Index).Text = IIf(Not IsNumeric(txtInteres(Index).Text), 0, txtInteres(Index).Text)
  txtInteres(Index).Text = FormatNumber(CDec(txtInteres(Index).Text), 2)
End Sub
Private Sub txtNroHoja_GotFocus()
  gdl_Procedure.MarcaGet txtNroHoja
End Sub
Private Sub txtNroHoja_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_GotFocus()
  gdl_Procedure.MarcaGet txtRemunera
End Sub
Private Sub txtRemunera_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtRemunera_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera, "CP")
End Sub
Private Sub txtRetencion_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtRetencion(Index)
End Sub
Private Sub txtRetencion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 5
End Sub
Private Sub txtRetencion_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRetencion_LostFocus(Index As Integer)
  lblHelp(Index + 5) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(Index), "CP")
End Sub
