VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fAbcAsistencia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9165
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   8550
   Icon            =   "abcasistencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   8550
   Begin TabDlg.SSTab tabRegister 
      Height          =   7950
      Left            =   75
      TabIndex        =   179
      Top             =   600
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   14023
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
      TabPicture(0)   =   "abcasistencia.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNombre"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCuadro(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCuadro(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraRegister"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin Threed.SSFrame fraRegister 
         Height          =   5070
         Left            =   120
         TabIndex        =   39
         Top             =   2430
         Width           =   7335
         _Version        =   65536
         _ExtentX        =   12938
         _ExtentY        =   8943
         _StockProps     =   14
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
         Begin MSComctlLib.TabStrip tasRegister 
            Height          =   390
            Left            =   120
            TabIndex        =   40
            Top             =   165
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   688
            Style           =   2
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   4
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Vacaciones"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Licencia"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Liquidación"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Permisos"
                  ImageVarType    =   2
               EndProperty
            EndProperty
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
         Begin Threed.SSFrame fraFicha 
            Height          =   4515
            Index           =   0
            Left            =   -9000
            TabIndex        =   41
            Top             =   480
            Width           =   7065
            _Version        =   65536
            _ExtentX        =   12462
            _ExtentY        =   7964
            _StockProps     =   14
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
            Begin VB.TextBox txtVacaciones 
               ForeColor       =   &H00000000&
               Height          =   280
               Index           =   1
               Left            =   1440
               TabIndex        =   54
               Top             =   2310
               Width           =   660
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   0
               Left            =   1440
               MaxLength       =   8
               TabIndex        =   50
               Top             =   840
               Width           =   660
            End
            Begin VB.TextBox txtVacaciones 
               ForeColor       =   &H00000000&
               Height          =   280
               Index           =   0
               Left            =   1440
               TabIndex        =   44
               Top             =   540
               Width           =   660
            End
            Begin MSMask.MaskEdBox mskFisVacacion 
               Height          =   285
               Index           =   0
               Left            =   5625
               TabIndex        =   65
               Top             =   2625
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFisVacacion 
               Height          =   285
               Index           =   1
               Left            =   5625
               TabIndex        =   67
               Top             =   2925
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskPerVacacion 
               Height          =   285
               Index           =   1
               Left            =   5655
               TabIndex        =   69
               Top             =   3405
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   9
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
            Begin MSMask.MaskEdBox mskFisVacacion 
               Height          =   285
               Index           =   2
               Left            =   5655
               TabIndex        =   71
               Top             =   3705
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFisVacacion 
               Height          =   285
               Index           =   3
               Left            =   5655
               TabIndex        =   73
               Top             =   4005
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskPerVacacion 
               Height          =   285
               Index           =   0
               Left            =   5625
               TabIndex        =   63
               Top             =   2325
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFecVacacion 
               Height          =   285
               Index           =   0
               Left            =   1440
               TabIndex        =   46
               Top             =   1140
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecVacacion 
               Height          =   285
               Index           =   1
               Left            =   1440
               TabIndex        =   48
               Top             =   1440
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   0
               Left            =   2160
               TabIndex        =   180
               Top             =   840
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin MSMask.MaskEdBox mskFisVacacion 
               Height          =   285
               Index           =   4
               Left            =   1440
               TabIndex        =   58
               Top             =   2910
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFisVacacion 
               Height          =   285
               Index           =   5
               Left            =   1440
               TabIndex        =   60
               Top             =   3210
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskPerVacacion 
               Height          =   285
               Index           =   2
               Left            =   1440
               TabIndex        =   56
               Top             =   2610
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin Threed.SSCheck chkAdeVacacion 
               Height          =   255
               Left            =   4545
               TabIndex        =   42
               Top             =   150
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Vacaciones Adelantadas"
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
               Caption         =   "Venta Vacaciones : "
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
               Height          =   210
               Index           =   24
               Left            =   225
               TabIndex        =   52
               Top             =   1965
               Width           =   1845
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               BorderStyle     =   6  'Inside Solid
               Height          =   1410
               Index           =   1
               Left            =   165
               Shape           =   4  'Rounded Rectangle
               Top             =   2205
               Width           =   2505
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   28
               Left            =   225
               TabIndex        =   59
               Top             =   3255
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Periodo :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   26
               Left            =   225
               TabIndex        =   55
               Top             =   2640
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   27
               Left            =   225
               TabIndex        =   57
               Top             =   2940
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Días :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   25
               Left            =   225
               TabIndex        =   53
               Top             =   2340
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Motivo : "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   21
               Left            =   225
               TabIndex        =   49
               Top             =   870
               Width           =   1140
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
               Left            =   2520
               TabIndex        =   51
               Top             =   870
               Width           =   195
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1380
               Index           =   0
               Left            =   165
               Shape           =   4  'Rounded Rectangle
               Top             =   450
               Width           =   6735
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   23
               Left            =   225
               TabIndex        =   47
               Top             =   1470
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   22
               Left            =   225
               TabIndex        =   45
               Top             =   1170
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Días :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   20
               Left            =   225
               TabIndex        =   43
               Top             =   570
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Caption         =   "Vacaciones Fisicas :"
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
               Height          =   210
               Index           =   29
               Left            =   4410
               TabIndex        =   61
               Top             =   1965
               Width           =   1845
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               BorderStyle     =   6  'Inside Solid
               Height          =   1095
               Index           =   3
               Left            =   4365
               Shape           =   4  'Rounded Rectangle
               Top             =   3285
               Width           =   2505
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               BorderStyle     =   6  'Inside Solid
               Height          =   1095
               Index           =   2
               Left            =   4335
               Shape           =   4  'Rounded Rectangle
               Top             =   2205
               Width           =   2505
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio 2 :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   34
               Left            =   4440
               TabIndex        =   70
               Top             =   3735
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final 2 :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   35
               Left            =   4440
               TabIndex        =   72
               Top             =   4035
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final 1 :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   32
               Left            =   4410
               TabIndex        =   66
               Top             =   2955
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Periodo 2 :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   33
               Left            =   4440
               TabIndex        =   68
               Top             =   3435
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Periodo 1 :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   30
               Left            =   4410
               TabIndex        =   62
               Top             =   2355
               Width           =   1140
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio 1 :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   31
               Left            =   4410
               TabIndex        =   64
               Top             =   2655
               Width           =   1140
            End
         End
         Begin Threed.SSFrame fraFicha 
            Height          =   4515
            Index           =   2
            Left            =   -20000
            TabIndex        =   115
            Top             =   480
            Width           =   7065
            _Version        =   65536
            _ExtentX        =   12462
            _ExtentY        =   7964
            _StockProps     =   14
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
            Begin VB.TextBox txtDiasLiqVacacion 
               ForeColor       =   &H00000000&
               Height          =   280
               Left            =   5145
               TabIndex        =   160
               Top             =   435
               Width           =   660
            End
            Begin VB.TextBox txtDiasLiquidacion 
               ForeColor       =   &H00800000&
               Height          =   280
               Left            =   1890
               TabIndex        =   119
               Top             =   750
               Width           =   660
            End
            Begin VB.TextBox txtDiasLiqGratifica 
               Height          =   280
               Left            =   1890
               TabIndex        =   121
               Top             =   1065
               Width           =   660
            End
            Begin VB.TextBox txtObservacion 
               ForeColor       =   &H00000000&
               Height          =   280
               Left            =   1890
               TabIndex        =   123
               Top             =   1395
               Width           =   4365
            End
            Begin MSMask.MaskEdBox mskFecVacaLiqui 
               Height          =   285
               Index           =   0
               Left            =   5145
               TabIndex        =   162
               Top             =   750
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecVacaLiqui 
               Height          =   285
               Index           =   1
               Left            =   5145
               TabIndex        =   164
               Top             =   1065
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFechaCese 
               Height          =   285
               Left            =   1890
               TabIndex        =   117
               Top             =   435
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin Threed.SSCheck chkLiquidaCalifica 
               Height          =   255
               Left            =   1890
               TabIndex        =   124
               Top             =   1725
               Width           =   2430
               _Version        =   65536
               _ExtentX        =   4286
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Desempeño no calificado"
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
            Begin Threed.SSFrame fraCuadro 
               Height          =   2130
               Index           =   2
               Left            =   105
               TabIndex        =   125
               Top             =   2295
               Width           =   6840
               _Version        =   65536
               _ExtentX        =   12065
               _ExtentY        =   3757
               _StockProps     =   14
               Caption         =   " Vacaciones Vencidas "
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
               Begin VB.TextBox txtVacaVencida 
                  ForeColor       =   &H00000000&
                  Height          =   280
                  Left            =   1785
                  TabIndex        =   127
                  Top             =   390
                  Width           =   660
               End
               Begin MSMask.MaskEdBox mskFecVacaVen 
                  Height          =   285
                  Index           =   0
                  Left            =   1785
                  TabIndex        =   131
                  Top             =   1200
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
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
               Begin MSMask.MaskEdBox mskFecVacaVen 
                  Height          =   285
                  Index           =   1
                  Left            =   1785
                  TabIndex        =   133
                  Top             =   1500
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
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
               Begin MSMask.MaskEdBox mskPerVacacion 
                  Height          =   285
                  Index           =   4
                  Left            =   5040
                  TabIndex        =   135
                  Top             =   900
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   9
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
               Begin MSMask.MaskEdBox mskFecVacaVen 
                  Height          =   285
                  Index           =   2
                  Left            =   5040
                  TabIndex        =   137
                  Top             =   1200
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
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
               Begin MSMask.MaskEdBox mskFecVacaVen 
                  Height          =   285
                  Index           =   3
                  Left            =   5040
                  TabIndex        =   139
                  Top             =   1500
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
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
               Begin MSMask.MaskEdBox mskPerVacacion 
                  Height          =   285
                  Index           =   3
                  Left            =   1785
                  TabIndex        =   129
                  Top             =   900
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   9
                  PromptChar      =   "_"
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Días :"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   61
                  Left            =   390
                  TabIndex        =   126
                  Top             =   420
                  Width           =   1365
               End
               Begin VB.Shape shpCuadro 
                  BorderColor     =   &H00C00000&
                  BorderStyle     =   6  'Inside Solid
                  Height          =   1155
                  Index           =   9
                  Left            =   120
                  Shape           =   4  'Rounded Rectangle
                  Top             =   765
                  Width           =   6630
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fecha Inicio 2 :"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   66
                  Left            =   3645
                  TabIndex        =   136
                  Top             =   1245
                  Width           =   1320
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fecha Final 2 :"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   67
                  Left            =   3645
                  TabIndex        =   138
                  Top             =   1545
                  Width           =   1320
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fecha Final 1 :"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   64
                  Left            =   390
                  TabIndex        =   132
                  Top             =   1545
                  Width           =   1320
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Periodo 2 :"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   65
                  Left            =   3645
                  TabIndex        =   134
                  Top             =   945
                  Width           =   1320
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Periodo 1 :"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   62
                  Left            =   390
                  TabIndex        =   128
                  Top             =   945
                  Width           =   1320
               End
               Begin VB.Label lblDato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fecha Inicio 1 :"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   63
                  Left            =   390
                  TabIndex        =   130
                  Top             =   1245
                  Width           =   1320
               End
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               BorderStyle     =   6  'Inside Solid
               Height          =   1785
               Index           =   8
               Left            =   90
               Shape           =   4  'Rounded Rectangle
               Top             =   300
               Width           =   6870
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Vac. Pendientes :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   58
               Left            =   3750
               TabIndex        =   159
               Top             =   480
               Width           =   1320
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Vac. Fin :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   60
               Left            =   3750
               TabIndex        =   163
               Top             =   1110
               Width           =   1320
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Vac. Ini :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   59
               Left            =   3750
               TabIndex        =   161
               Top             =   795
               Width           =   1320
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Dias Liquidación :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   55
               Left            =   450
               TabIndex        =   118
               Top             =   795
               Width           =   1365
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Dias Gratificación :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   56
               Left            =   450
               TabIndex        =   120
               Top             =   1110
               Width           =   1365
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Cese :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   54
               Left            =   450
               TabIndex        =   116
               Top             =   480
               Width           =   1365
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Observación :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   57
               Left            =   450
               TabIndex        =   122
               Top             =   1440
               Width           =   1365
            End
         End
         Begin Threed.SSFrame fraFicha 
            Height          =   4515
            Index           =   1
            Left            =   -10000
            TabIndex        =   74
            Top             =   480
            Width           =   7065
            _Version        =   65536
            _ExtentX        =   12462
            _ExtentY        =   7964
            _StockProps     =   14
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
            Begin VB.TextBox txtCertificado 
               Height          =   280
               Index           =   1
               Left            =   2385
               TabIndex        =   105
               Top             =   2295
               Width           =   1035
            End
            Begin VB.TextBox txtCertificado 
               Height          =   280
               Index           =   0
               Left            =   2385
               TabIndex        =   85
               Top             =   660
               Width           =   1035
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   4
               Left            =   4710
               TabIndex        =   109
               Top             =   2295
               Width           =   375
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   3
               Left            =   1380
               TabIndex        =   98
               Top             =   2310
               Width           =   375
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   1
               Left            =   1380
               TabIndex        =   78
               Top             =   660
               Width           =   375
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   2
               Left            =   4710
               TabIndex        =   89
               Top             =   660
               Width           =   375
            End
            Begin VB.TextBox txtEnfermedad 
               Height          =   280
               Left            =   1380
               TabIndex        =   96
               Top             =   1995
               Width           =   660
            End
            Begin VB.TextBox txtLicencia 
               ForeColor       =   &H00000000&
               Height          =   280
               Left            =   4710
               TabIndex        =   107
               Top             =   1995
               Width           =   660
            End
            Begin VB.TextBox txtPrePostNatal 
               ForeColor       =   &H00800000&
               Height          =   280
               Left            =   1380
               TabIndex        =   76
               Top             =   345
               Width           =   660
            End
            Begin VB.TextBox txtAccidente 
               ForeColor       =   &H00000000&
               Height          =   280
               Left            =   4710
               TabIndex        =   87
               Top             =   345
               Width           =   660
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   2
               Left            =   5145
               TabIndex        =   182
               Top             =   660
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
               Left            =   1815
               TabIndex        =   181
               Top             =   660
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin MSMask.MaskEdBox mskFecPreNatal 
               Height          =   285
               Index           =   0
               Left            =   1380
               TabIndex        =   81
               Top             =   1260
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecPreNatal 
               Height          =   285
               Index           =   1
               Left            =   1380
               TabIndex        =   83
               Top             =   1560
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecAccidente 
               Height          =   285
               Index           =   0
               Left            =   4770
               TabIndex        =   92
               Top             =   1260
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecAccidente 
               Height          =   285
               Index           =   1
               Left            =   4770
               TabIndex        =   94
               Top             =   1560
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   3
               Left            =   1815
               TabIndex        =   183
               Top             =   2295
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin MSMask.MaskEdBox mskFecEnfermedad 
               Height          =   285
               Index           =   0
               Left            =   1380
               TabIndex        =   101
               Top             =   2895
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecEnfermedad 
               Height          =   285
               Index           =   1
               Left            =   1380
               TabIndex        =   103
               Top             =   3195
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   4
               Left            =   5145
               TabIndex        =   184
               Top             =   2295
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin MSMask.MaskEdBox mskFecLicencia 
               Height          =   285
               Index           =   0
               Left            =   4710
               TabIndex        =   112
               Top             =   2895
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecLicencia 
               Height          =   285
               Index           =   1
               Left            =   4710
               TabIndex        =   114
               Top             =   3195
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Nro CITT :"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   43
               Left            =   2235
               TabIndex        =   104
               Top             =   1995
               Width           =   1095
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1665
               Index           =   7
               Left            =   3525
               Shape           =   4  'Rounded Rectangle
               Top             =   1920
               Width           =   3435
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1665
               Index           =   5
               Left            =   3525
               Shape           =   4  'Rounded Rectangle
               Top             =   270
               Width           =   3435
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1665
               Index           =   6
               Left            =   105
               Shape           =   4  'Rounded Rectangle
               Top             =   1920
               Width           =   3435
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Nro CITT :"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   38
               Left            =   2235
               TabIndex        =   84
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   52
               Left            =   3555
               TabIndex        =   111
               Top             =   2940
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   53
               Left            =   3555
               TabIndex        =   113
               Top             =   3240
               Width           =   1095
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
               Left            =   4050
               TabIndex        =   110
               Top             =   2640
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Motivo : "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   51
               Left            =   3555
               TabIndex        =   108
               Top             =   2310
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   44
               Left            =   195
               TabIndex        =   100
               Top             =   2940
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   45
               Left            =   195
               TabIndex        =   102
               Top             =   3240
               Width           =   1095
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
               Left            =   630
               TabIndex        =   99
               Top             =   2640
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Motivo : "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   42
               Left            =   195
               TabIndex        =   97
               Top             =   2325
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   48
               Left            =   3615
               TabIndex        =   91
               Top             =   1305
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   49
               Left            =   3615
               TabIndex        =   93
               Top             =   1605
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   39
               Left            =   195
               TabIndex        =   80
               Top             =   1305
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   40
               Left            =   195
               TabIndex        =   82
               Top             =   1605
               Width           =   1095
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1665
               Index           =   4
               Left            =   105
               Shape           =   4  'Rounded Rectangle
               Top             =   270
               Width           =   3435
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
               Left            =   630
               TabIndex        =   79
               Top             =   960
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Motivo :"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   37
               Left            =   195
               TabIndex        =   77
               Top             =   690
               Width           =   1095
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
               Left            =   4050
               TabIndex        =   90
               Top             =   960
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Motivo : "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   47
               Left            =   3555
               TabIndex        =   88
               Top             =   690
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Enfermedad :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   41
               Left            =   195
               TabIndex        =   95
               Top             =   2010
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Licencia :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   50
               Left            =   3555
               TabIndex        =   106
               Top             =   2010
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Natalidad :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   36
               Left            =   195
               TabIndex        =   75
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Accidente :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   46
               Left            =   3555
               TabIndex        =   86
               Top             =   360
               Width           =   1095
            End
         End
         Begin Threed.SSFrame fraFicha 
            Height          =   4515
            Index           =   3
            Left            =   -30000
            TabIndex        =   140
            Top             =   480
            Width           =   7065
            _Version        =   65536
            _ExtentX        =   12462
            _ExtentY        =   7964
            _StockProps     =   14
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
            Begin VB.TextBox txtPaternidad 
               ForeColor       =   &H00800000&
               Height          =   280
               Left            =   1380
               TabIndex        =   142
               Top             =   540
               Width           =   660
            End
            Begin VB.TextBox txtFallece 
               Height          =   280
               Left            =   1380
               TabIndex        =   151
               Top             =   2190
               Width           =   660
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   5
               Left            =   1380
               TabIndex        =   144
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox txtMotivo 
               Height          =   280
               Index           =   6
               Left            =   1380
               TabIndex        =   153
               Top             =   2490
               Width           =   375
            End
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   5
               Left            =   1815
               TabIndex        =   185
               Top             =   840
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin MSMask.MaskEdBox mskFecPaternidad 
               Height          =   285
               Index           =   0
               Left            =   1380
               TabIndex        =   147
               Top             =   1140
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecPaternidad 
               Height          =   285
               Index           =   1
               Left            =   1380
               TabIndex        =   149
               Top             =   1440
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin Threed.SSCommand cmdHelp 
               Height          =   285
               Index           =   6
               Left            =   1815
               TabIndex        =   186
               Top             =   2475
               Width           =   285
               _Version        =   65536
               _ExtentX        =   494
               _ExtentY        =   494
               _StockProps     =   78
               Caption         =   "..."
               Enabled         =   0   'False
            End
            Begin MSMask.MaskEdBox mskFecFallece 
               Height          =   285
               Index           =   0
               Left            =   1380
               TabIndex        =   156
               Top             =   2790
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox mskFecFallece 
               Height          =   285
               Index           =   1
               Left            =   1380
               TabIndex        =   158
               Top             =   3090
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
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
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Paternidad :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   68
               Left            =   195
               TabIndex        =   141
               Top             =   585
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fallecimiento :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   72
               Left            =   195
               TabIndex        =   150
               Top             =   2235
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Motivo :"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   69
               Left            =   195
               TabIndex        =   143
               Top             =   885
               Width           =   1095
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
               Left            =   2175
               TabIndex        =   145
               Top             =   885
               Width           =   195
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1365
               Index           =   13
               Left            =   105
               Shape           =   4  'Rounded Rectangle
               Top             =   465
               Width           =   6795
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   71
               Left            =   195
               TabIndex        =   148
               Top             =   1485
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   70
               Left            =   195
               TabIndex        =   146
               Top             =   1185
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Motivo : "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   73
               Left            =   195
               TabIndex        =   152
               Top             =   2535
               Width           =   1095
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
               Left            =   2175
               TabIndex        =   154
               Top             =   2520
               Width           =   195
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Final :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   75
               Left            =   195
               TabIndex        =   157
               Top             =   3135
               Width           =   1095
            End
            Begin VB.Label lblDato 
               Alignment       =   1  'Right Justify
               Caption         =   "Fecha Inicio :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   74
               Left            =   195
               TabIndex        =   155
               Top             =   2835
               Width           =   1095
            End
            Begin VB.Shape shpCuadro 
               BorderColor     =   &H00C00000&
               Height          =   1365
               Index           =   12
               Left            =   105
               Shape           =   4  'Rounded Rectangle
               Top             =   2115
               Width           =   6795
            End
         End
      End
      Begin Threed.SSFrame fraCuadro 
         Height          =   1950
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   480
         Width           =   3675
         _Version        =   65536
         _ExtentX        =   6470
         _ExtentY        =   3440
         _StockProps     =   14
         Caption         =   " Días "
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
         Begin VB.TextBox txtDiaLibre 
            ForeColor       =   &H00000000&
            Height          =   280
            Left            =   2910
            TabIndex        =   19
            Top             =   1215
            Width           =   660
         End
         Begin VB.TextBox txtDiaParcial 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   7
            Top             =   915
            Width           =   660
         End
         Begin VB.TextBox txtDiaSuspension 
            ForeColor       =   &H00000000&
            Height          =   280
            Left            =   2910
            TabIndex        =   17
            Top             =   915
            Width           =   660
         End
         Begin VB.TextBox txtDiaSemanal 
            ForeColor       =   &H00000000&
            Height          =   280
            Left            =   2910
            TabIndex        =   13
            Top             =   315
            Width           =   660
         End
         Begin VB.TextBox txtDiaFeriado 
            ForeColor       =   &H00000000&
            Height          =   280
            Left            =   1095
            TabIndex        =   11
            Top             =   1515
            Width           =   660
         End
         Begin VB.TextBox txtDiaTrabajo 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   3
            Top             =   315
            Width           =   660
         End
         Begin VB.TextBox txtFaltas 
            Height          =   280
            Left            =   2910
            TabIndex        =   15
            Top             =   615
            Width           =   660
         End
         Begin VB.TextBox txtDiaLaboral 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   9
            Top             =   1215
            Width           =   660
         End
         Begin VB.TextBox txtDiaMedio 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   5
            Top             =   615
            Width           =   660
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Libres :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   8
            Left            =   1905
            TabIndex        =   18
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Suspensión :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   7
            Left            =   1905
            TabIndex        =   16
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Trab. DSO :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   5
            Left            =   1905
            TabIndex        =   12
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Feriados :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   10
            Top             =   1560
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Trabajados :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   2
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Faltas :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   6
            Left            =   1905
            TabIndex        =   14
            Top             =   660
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Laborados :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   8
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Half-Time :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   660
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Part-Time :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   960
            Width           =   945
         End
      End
      Begin Threed.SSFrame fraCuadro 
         Height          =   1950
         Index           =   1
         Left            =   3810
         TabIndex        =   20
         Top             =   480
         Width           =   3675
         _Version        =   65536
         _ExtentX        =   6470
         _ExtentY        =   3440
         _StockProps     =   14
         Caption         =   " Horas "
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
         Begin VB.TextBox txtTardanzas 
            Height          =   280
            Left            =   2910
            TabIndex        =   38
            Top             =   1215
            Width           =   660
         End
         Begin VB.TextBox txtHoraParcial 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   26
            Top             =   915
            Width           =   660
         End
         Begin VB.TextBox txtHoraMedio 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   24
            Top             =   615
            Width           =   660
         End
         Begin VB.TextBox txtHrExtraDoble 
            ForeColor       =   &H00000000&
            Height          =   280
            Left            =   1095
            TabIndex        =   30
            Top             =   1515
            Width           =   660
         End
         Begin VB.TextBox txtHrExtraSimple 
            Height          =   280
            Left            =   1095
            TabIndex        =   28
            Top             =   1215
            Width           =   660
         End
         Begin VB.TextBox txtHoraNormal 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   1095
            TabIndex        =   22
            Top             =   315
            Width           =   660
         End
         Begin VB.TextBox txtHrEspecial 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   2910
            TabIndex        =   32
            Top             =   315
            Width           =   660
         End
         Begin VB.TextBox txtOpcional 
            ForeColor       =   &H00000000&
            Height          =   280
            Left            =   2910
            TabIndex        =   36
            Top             =   915
            Width           =   660
         End
         Begin VB.TextBox txtHrNocturno 
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   2910
            TabIndex        =   34
            Top             =   615
            Width           =   660
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Tardanzas :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   18
            Left            =   1905
            TabIndex        =   37
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Part-Time :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   90
            TabIndex        =   25
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Half-Time :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   23
            Top             =   660
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "H.E. Dobles :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   29
            Top             =   1560
            Width           =   950
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "H.E. 25% :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   90
            TabIndex        =   27
            Top             =   1260
            Width           =   950
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Normales :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   21
            Top             =   360
            Width           =   950
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "H.E. 35% :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   15
            Left            =   1905
            TabIndex        =   31
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Opcional :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   17
            Left            =   1905
            TabIndex        =   35
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Nocturno :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   16
            Left            =   1905
            TabIndex        =   33
            Top             =   660
            Width           =   945
         End
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre :"
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
         Height          =   300
         Left            =   200
         TabIndex        =   0
         Top             =   165
         Width           =   840
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   165
      Top             =   0
      Width           =   8550
      _Version        =   65536
      _ExtentX        =   15081
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
         Left            =   7695
         TabIndex        =   166
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
         Picture         =   "abcasistencia.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   7305
         TabIndex        =   167
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
         Picture         =   "abcasistencia.frx":0044
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
         Left            =   675
         TabIndex        =   168
         Top             =   120
         Width           =   6075
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   169
      Top             =   8655
      Width           =   8550
      _Version        =   65536
      _ExtentX        =   15081
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
         Left            =   5505
         TabIndex        =   170
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
         Picture         =   "abcasistencia.frx":0060
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   5115
         TabIndex        =   171
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
         Picture         =   "abcasistencia.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   3405
         TabIndex        =   172
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
         Picture         =   "abcasistencia.frx":0098
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   3015
         TabIndex        =   173
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
         Picture         =   "abcasistencia.frx":00B4
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   7950
      Index           =   0
      Left            =   7755
      TabIndex        =   174
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   14023
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
         TabIndex        =   175
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
         TabIndex        =   176
         Tag             =   "0"
         Top             =   810
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483631
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
         Picture         =   "abcasistencia.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   177
         Tag             =   "0"
         Top             =   1440
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483631
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
         Picture         =   "abcasistencia.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   178
         Tag             =   "0"
         Top             =   2040
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483631
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
         Picture         =   "abcasistencia.frx":0108
      End
   End
   Begin MSMask.MaskEdBox mskFecVacacion 
      Height          =   300
      Index           =   5
      Left            =   1275
      TabIndex        =   187
      Top             =   825
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2475
      TabIndex        =   188
      Top             =   210
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
Attribute VB_Name = "fAbcAsistencia"
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
Private n_IndexTabs As Integer                          ' indicce la pestaña del tab control
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private l_Inicializa As Boolean                         ' Flag de registro por rango
Private n_JornadaLaboral As Double                      ' Dias de jornada laboral
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  For n_Index = 0 To 6: cmdHelp(n_Index).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)

End Sub
Sub ShowScreen()
    
  lblNombre.Caption = "Personal : Inicialización de Asistencia de Personal "
  n_JornadaLaboral = pn_HoroLaboraxDia
  If Not l_Inicializa Then
    s_Sql = "SELECT diatrabajo, diamediotm, diaparcial, dialaboral, horanormal, horamediotm, horaparcial, horatipo1, "
    s_Sql = s_Sql & "horatipo2, horatipo3, horatipo4, diafalta, tardanza, diaprepostnatal, accidente, diavacaciones, enfermedad, licencia, "
    s_Sql = s_Sql & "diaferiado, diatradesemanal, diasuspension, dialibre, permisos, opcional, fechainivacacion, fechafinvacacion, dialiquidacion, "
    s_Sql = s_Sql & "pdovaca1, fechainivaca1, fechafinvaca1, pdovaca2, fechainivaca2, fechafinvaca2, "
    s_Sql = s_Sql & "liquidavacacion, diagratificacion, fechacese, fechainiliqvaca, fechafinliqvaca, observacion, liqnocalifica, "
    s_Sql = s_Sql & "indvacadelanta, diavacaventa, pdovaca3, fechainivaca3, fechafinvaca3, codmdi_vacac, "
    s_Sql = s_Sql & "diavacavencida, pdovaca4, fechainivaca4, fechafinvaca4, pdovaca5, fechainivaca5, fechafinvaca5, "
    s_Sql = s_Sql & "codmdi_natal, fechaini_natal, fechafin_natal, numecitt_natal, "
    s_Sql = s_Sql & "codmdi_accid, fechaini_accid, fechafin_accid, "
    s_Sql = s_Sql & "codmdi_enfer, fechaini_enfer, fechafin_enfer, numecitt_enfer, "
    s_Sql = s_Sql & "codmdi_licen, fechaini_licen, fechafin_licen, "
    s_Sql = s_Sql & "diapaternidad, codmdi_pater, fechaini_pater, fechafin_pater, "
    s_Sql = s_Sql & "diafallecefam, codmdi_falle, fechaini_falle, fechafin_falle "
    s_Sql = s_Sql & "FROM plasistencia "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpdo='" & Trim(o_SelAsistencia.txtPeriodo.Text) & "' "
    s_Sql = s_Sql & "AND codpsn='" & Trim(o_SelAsistencia.dcaRegistro.Recordset!codpsn) & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Cambio el tipo de presentación
    Me.Tag = IIf((porstRecordset.RecordCount = 0), s_MdoData_Ins, Me.Tag)
    lblNombre = "Personal : " & Trim(o_SelAsistencia.dcaRegistro.Recordset!apepaterno) & " " & Trim(o_SelAsistencia.dcaRegistro.Recordset!apematerno) & "; " & Trim(o_SelAsistencia.dcaRegistro.Recordset!nombres) & "  "
    n_JornadaLaboral = IIf(CDec(o_SelAsistencia.dcaRegistro.Recordset!jornadalaboral) > 0, CDec(o_SelAsistencia.dcaRegistro.Recordset!jornadalaboral), n_JornadaLaboral)
  End If
    
  ' Presenta Botones y Controles
  EnabledBotons
  
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    ' Generales
    gdl_Procedure.EditText "AT", txtDiaTrabajo, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaMedio, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaParcial, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaLaboral, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFaltas, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtTardanzas, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaFeriado, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaSemanal, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaSuspension, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaLibre, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHoraNormal, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHoraMedio, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHoraParcial, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrExtraSimple, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrExtraDoble, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrEspecial, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrNocturno, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtOpcional, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    ' Inicial ficha
    gdl_Procedure.EditOptionCheck "AT", chkAdeVacacion, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtVacaciones(0), 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(0), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecVacacion(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacacion(1), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskPerVacacion(0), "", Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(1), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskPerVacacion(1), "", Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(2), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(3), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtVacaciones(1), 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditMask "AT", mskPerVacacion(2), "", Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(4), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(5), "", Me.Tag, True, "##/##/####"
    ' Primera ficha
    gdl_Procedure.EditText "AT", txtPrePostNatal, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(1), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecPreNatal(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecPreNatal(1), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtCertificado(0), "", Me.Tag, False, 2, vbLeftJustify
    
    gdl_Procedure.EditText "AT", txtAccidente, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(2), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecAccidente(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecAccidente(1), "", Me.Tag, True, "##/##/####"
    
    gdl_Procedure.EditText "AT", txtEnfermedad, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(3), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecEnfermedad(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecEnfermedad(1), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtCertificado(1), "", Me.Tag, False, 2, vbLeftJustify
    
    gdl_Procedure.EditText "AT", txtLicencia, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(4), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecLicencia(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecLicencia(1), "", Me.Tag, True, "##/##/####"
    ' Segunda ficha
    gdl_Procedure.EditMask "AT", mskFechaCese, "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtDiasLiquidacion, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiasLiqGratifica, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiasLiqVacacion, 0, Me.Tag, False, 6, vbRightJustify, "#0.000"
    gdl_Procedure.EditMask "AT", mskFecVacaLiqui(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacaLiqui(1), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtObservacion, "", Me.Tag, False, 60, vbLeftJustify
    gdl_Procedure.EditOptionCheck "AT", chkLiquidaCalifica, False, Me.Tag, True
    
    gdl_Procedure.EditText "AT", txtVacaVencida, 0, Me.Tag, False, 6, vbRightJustify, "#0.000"
    gdl_Procedure.EditMask "AT", mskPerVacacion(3), "", Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(1), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskPerVacacion(4), "", Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(2), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(3), "", Me.Tag, True, "##/##/####"
    ' Tercera ficha
    gdl_Procedure.EditText "AT", txtpaternidad, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(5), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecPaternidad(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecPaternidad(1), "", Me.Tag, True, "##/##/####"
    
    gdl_Procedure.EditText "AT", txtFallece, 0, Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(6), "", Me.Tag, False, 2, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecFallece(0), "", Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecFallece(1), "", Me.Tag, True, "##/##/####"
  Else
    ' Generales
    gdl_Procedure.EditText "AT", txtDiaTrabajo, FormatNumber(porstRecordset!diatrabajo, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaMedio, FormatNumber(porstRecordset!diamediotm, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaParcial, FormatNumber(porstRecordset!diaparcial, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaLaboral, FormatNumber(porstRecordset!dialaboral, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFaltas, FormatNumber(porstRecordset!diafalta, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtTardanzas, FormatNumber(porstRecordset!tardanza, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaFeriado, FormatNumber(porstRecordset!diaferiado, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaSemanal, FormatNumber(porstRecordset!diatradesemanal, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaSuspension, FormatNumber(porstRecordset!diasuspension, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiaLibre, FormatNumber(porstRecordset!dialibre, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHoraNormal, FormatNumber(porstRecordset!horanormal, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHoraMedio, FormatNumber(porstRecordset!horamediotm, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHoraParcial, FormatNumber(porstRecordset!horaparcial, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrExtraSimple, FormatNumber(porstRecordset!horatipo1, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrExtraDoble, FormatNumber(porstRecordset!horatipo2, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrEspecial, FormatNumber(porstRecordset!horatipo3, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtHrNocturno, FormatNumber(porstRecordset!horatipo4, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtOpcional, FormatNumber(porstRecordset!opcional, 2), Me.Tag, False, 6, vbRightJustify
    ' Inicial ficha
    gdl_Procedure.EditOptionCheck "AT", chkAdeVacacion, (CInt(porstRecordset!indvacadelanta) = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtVacaciones(0), FormatNumber(porstRecordset!diavacaciones, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(0), gdl_Funcion.aTexto(porstRecordset!codmdi_vacac), Me.Tag, False, porstRecordset!codmdi_vacac.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecVacacion(0), IIf(IsNull(porstRecordset!fechainivacacion), "", porstRecordset!fechainivacacion), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacacion(1), IIf(IsNull(porstRecordset!fechafinvacacion), "", porstRecordset!fechafinvacacion), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskPerVacacion(0), Format(IIf(IsNull(porstRecordset!pdovaca1), "", porstRecordset!pdovaca1), "####-####"), Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(0), IIf(IsNull(porstRecordset!fechainivaca1), "", porstRecordset!fechainivaca1), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(1), IIf(IsNull(porstRecordset!fechafinvaca1), "", porstRecordset!fechafinvaca1), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskPerVacacion(1), Format(IIf(IsNull(porstRecordset!pdovaca2), "", porstRecordset!pdovaca2), "####-####"), Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(2), IIf(IsNull(porstRecordset!fechainivaca2), "", porstRecordset!fechainivaca2), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(3), IIf(IsNull(porstRecordset!fechafinvaca2), "", porstRecordset!fechafinvaca2), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtVacaciones(1), FormatNumber(porstRecordset!diavacaventa, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditMask "AT", mskPerVacacion(2), Format(IIf(IsNull(porstRecordset!pdovaca3), "", porstRecordset!pdovaca3), "####-####"), Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(4), IIf(IsNull(porstRecordset!fechainivaca3), "", porstRecordset!fechainivaca3), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFisVacacion(5), IIf(IsNull(porstRecordset!fechafinvaca3), "", porstRecordset!fechafinvaca3), Me.Tag, True, "##/##/####"
    ' Primera ficha
    gdl_Procedure.EditText "AT", txtPrePostNatal, FormatNumber(porstRecordset!diaprepostnatal, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(1), gdl_Funcion.aTexto(porstRecordset!codmdi_natal), Me.Tag, False, porstRecordset!codmdi_natal.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecPreNatal(0), IIf(IsNull(porstRecordset!fechaini_natal), "", porstRecordset!fechaini_natal), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecPreNatal(1), IIf(IsNull(porstRecordset!fechafin_natal), "", porstRecordset!fechafin_natal), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtCertificado(0), gdl_Funcion.aTexto(porstRecordset!numecitt_natal), Me.Tag, False, porstRecordset!numecitt_natal.DefinedSize, vbLeftJustify
    
    gdl_Procedure.EditText "AT", txtEnfermedad, FormatNumber(porstRecordset!enfermedad, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(3), gdl_Funcion.aTexto(porstRecordset!codmdi_enfer), Me.Tag, False, porstRecordset!codmdi_enfer.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecEnfermedad(0), IIf(IsNull(porstRecordset!fechaini_enfer), "", porstRecordset!fechaini_enfer), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecEnfermedad(1), IIf(IsNull(porstRecordset!fechafin_enfer), "", porstRecordset!fechafin_enfer), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtCertificado(1), gdl_Funcion.aTexto(porstRecordset!numecitt_enfer), Me.Tag, False, porstRecordset!numecitt_enfer.DefinedSize, vbLeftJustify
    
    gdl_Procedure.EditText "AT", txtAccidente, FormatNumber(porstRecordset!accidente, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(2), gdl_Funcion.aTexto(porstRecordset!codmdi_accid), Me.Tag, False, porstRecordset!codmdi_accid.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecAccidente(0), IIf(IsNull(porstRecordset!fechaini_accid), "", porstRecordset!fechaini_accid), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecAccidente(1), IIf(IsNull(porstRecordset!fechafin_accid), "", porstRecordset!fechafin_accid), Me.Tag, True, "##/##/####"
    
    gdl_Procedure.EditText "AT", txtLicencia, FormatNumber(porstRecordset!licencia, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(4), gdl_Funcion.aTexto(porstRecordset!codmdi_licen), Me.Tag, False, porstRecordset!codmdi_licen.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecLicencia(0), IIf(IsNull(porstRecordset!fechaini_licen), "", porstRecordset!fechaini_licen), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecLicencia(1), IIf(IsNull(porstRecordset!fechafin_licen), "", porstRecordset!fechafin_licen), Me.Tag, True, "##/##/####"
    
    ' Segunda ficha
    gdl_Procedure.EditMask "AT", mskFechaCese, IIf(IsNull(porstRecordset!fechacese), "", porstRecordset!fechacese), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtDiasLiquidacion, FormatNumber(porstRecordset!dialiquidacion, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiasLiqGratifica, FormatNumber(porstRecordset!diagratificacion, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtDiasLiqVacacion, FormatNumber(porstRecordset!liquidavacacion, 3), Me.Tag, False, 6, vbRightJustify, "#0.000"
    gdl_Procedure.EditMask "AT", mskFecVacaLiqui(0), IIf(IsNull(porstRecordset!fechainiliqvaca), "", porstRecordset!fechainiliqvaca), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacaLiqui(1), IIf(IsNull(porstRecordset!fechafinliqvaca), "", porstRecordset!fechafinliqvaca), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditText "AT", txtObservacion, gdl_Funcion.aTexto(porstRecordset!observacion), Me.Tag, False, porstRecordset!observacion.DefinedSize, vbLeftJustify
    gdl_Procedure.EditOptionCheck "AT", chkLiquidaCalifica, (CInt(porstRecordset!liqnocalifica) = s_Estado_Act), Me.Tag, True
  
    gdl_Procedure.EditText "AT", txtVacaVencida, FormatNumber(porstRecordset!diavacavencida, 3), Me.Tag, False, 6, vbRightJustify, "#0.000"
    gdl_Procedure.EditMask "AT", mskPerVacacion(3), Format(IIf(IsNull(porstRecordset!pdovaca4), "", porstRecordset!pdovaca4), "####-####"), Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(0), IIf(IsNull(porstRecordset!fechainivaca4), "", porstRecordset!fechainivaca4), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(1), IIf(IsNull(porstRecordset!fechafinvaca4), "", porstRecordset!fechafinvaca4), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskPerVacacion(4), Format(IIf(IsNull(porstRecordset!pdovaca5), "", porstRecordset!pdovaca5), "####-####"), Me.Tag, True, "####-####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(2), IIf(IsNull(porstRecordset!fechainivaca5), "", porstRecordset!fechainivaca5), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecVacaVen(3), IIf(IsNull(porstRecordset!fechafinvaca5), "", porstRecordset!fechafinvaca5), Me.Tag, True, "##/##/####"
  
    ' Tercera ficha
    gdl_Procedure.EditText "AT", txtpaternidad, FormatNumber(porstRecordset!diapaternidad, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(5), gdl_Funcion.aTexto(porstRecordset!codmdi_pater), Me.Tag, False, porstRecordset!codmdi_pater.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecPaternidad(0), IIf(IsNull(porstRecordset!fechaini_pater), "", porstRecordset!fechaini_pater), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecPaternidad(1), IIf(IsNull(porstRecordset!fechafin_pater), "", porstRecordset!fechafin_pater), Me.Tag, True, "##/##/####"
    
    gdl_Procedure.EditText "AT", txtFallece, FormatNumber(porstRecordset!diafallecefam, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtMotivo(6), gdl_Funcion.aTexto(porstRecordset!codmdi_falle), Me.Tag, False, porstRecordset!codmdi_falle.DefinedSize, vbLeftJustify
    gdl_Procedure.EditMask "AT", mskFecFallece(0), IIf(IsNull(porstRecordset!fechaini_falle), "", porstRecordset!fechaini_falle), Me.Tag, True, "##/##/####"
    gdl_Procedure.EditMask "AT", mskFecFallece(1), IIf(IsNull(porstRecordset!fechafin_falle), "", porstRecordset!fechafin_falle), Me.Tag, True, "##/##/####"
  End If
  For n_Index = 0 To 6
    lblHelp(n_Index).Caption = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtMotivo(n_Index).Text, "TS"), 30)
  Next n_Index
  txtDiaLaboral.Tag = txtDiaLaboral.Text
  Set porstRecordset = Nothing

End Sub
']
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Or Index = 2 Then
   txtDiaTrabajo.SetFocus
  End If
  If Index <> 1 Then Exit Sub
    
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim(o_SelAsistencia.dcaRegistro.Recordset!codpsn)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpdo", "codpsn")
    a_Valores = Array(ps_ClsPlanilla, Trim(o_SelAsistencia.txtPeriodo.Text), s_Registro)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plasistencia", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl o_SelAsistencia.dcaRegistro, o_SelAsistencia.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((o_SelAsistencia.dcaRegistro.Recordset.EOF And o_SelAsistencia.dcaRegistro.Recordset.BOF) Or o_SelAsistencia.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      o_SelAsistencia.dcaRegistro.Recordset.Find ("codpsn >= '" & s_Registro & "'")
      If o_SelAsistencia.dcaRegistro.Recordset.EOF Then o_SelAsistencia.dcaRegistro.Recordset.MoveLast
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
    
  If (Me.Tag = s_MdoData_Vis Or l_ExistRecord Or l_Inicializa) Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim nTop As Integer, nLeft As Integer
  
  s_SqlHelp = ""
  Select Case Index
   Case 0, 1, 2, 3, 4, 5, 6 ' motivo inasistencia
    tdbHelp.Columns(0).DataField = "codtsu": tdbHelp.Columns(1).DataField = "destsu"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("tsu", tdbHelp.Columns(0).DataField, "", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  nTop = tabRegister.Top + fraRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)) + IIf(Index < 3, fraFicha(1).Top, (fraFicha(1).Top / 2))
  nLeft = fraRegister.Left + cmdHelp(Index).Left + (cmdHelp(Index).Width / 2)
  nLeft = Choose(Index + 1, 3850, nLeft, 3000, nLeft, 3000, nLeft, nLeft)
  tdbHelp.Top = nTop
  tdbHelp.Left = nLeft
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: o_SelAsistencia.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not o_SelAsistencia.dcaRegistro.Recordset.BOF Then o_SelAsistencia.dcaRegistro.Recordset.MovePrevious
           If o_SelAsistencia.dcaRegistro.Recordset.BOF Then o_SelAsistencia.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not o_SelAsistencia.dcaRegistro.Recordset.EOF Then o_SelAsistencia.dcaRegistro.Recordset.MoveNext
           If o_SelAsistencia.dcaRegistro.Recordset.EOF Then o_SelAsistencia.dcaRegistro.Recordset.MoveLast
   Case 3: o_SelAsistencia.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_FechaHora As String, s_OldMessage As String
  Dim porstClone As ADODB.Recordset
  Dim nSecuencia As Long, nRegistroID As Long, nRegistroIX As Long
  Dim nIndexIni As Integer, nIndexFin As Integer
  Dim a_Vacaciones(), a_AcumVacacion()
  Dim nDiaVacacion As Long, nDAcuVacacion As Long
  
  ' FICHA INICIAL - Vacaciones
  ' Vacaciones global - información y fechas
  If chkAdeVacacion.Value And CDec(txtVacaciones(0).Text) <= 0 Then Beep: MsgBox "Debe Ingresar Días de Vacaciones Fisicas ", vbExclamation: tasRegister.Tabs(1).Selected = True: txtVacaciones(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) > 0 And txtMotivo(0).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(1).Selected = True: txtMotivo(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) > 0 And lblHelp(0).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(1).Selected = True: txtMotivo(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) > 0 And mskFecVacacion(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Vacaciones ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFecVacacion(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) > 0 And mskFecVacacion(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Vacaciones ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFecVacacion(1).SetFocus: Exit Sub
  If mskFecVacacion(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacacion(0), 1900) Then tasRegister.Tabs(1).Selected = True: mskFecVacacion(0).SetFocus: Exit Sub
  End If
  If mskFecVacacion(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacacion(1), 1900) Then tasRegister.Tabs(1).Selected = True: mskFecVacacion(1).SetFocus: Exit Sub
  End If
  
  ' Vacaciones Fisicas - información
  If CDec(txtVacaciones(0).Text) > 0 And Len(mskPerVacacion(0).ClipText) < 8 Then Beep: MsgBox "Debe Ingresar Periodo de Vacaciones Fisicas ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskPerVacacion(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) = 0 And Len(mskPerVacacion(0).ClipText) <> 0 Then Beep: MsgBox "Periodo de Vacaciones Fisicas no Valido", vbExclamation: tasRegister.Tabs(1).Selected = True: mskPerVacacion(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) > 0 And mskFisVacacion(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicial de Vacaciones Fisicas ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFisVacacion(0).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) > 0 And mskFisVacacion(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Vacaciones Fisicas ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFisVacacion(1).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) = 0 And Len(mskPerVacacion(1).ClipText) <> 0 Then Beep: MsgBox "Periodo de Vacaciones Fisicas no Valido", vbExclamation: tasRegister.Tabs(1).Selected = True: mskPerVacacion(1).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) = 0 And mskFisVacacion(2).ClipText <> "" Then Beep: MsgBox "Fecha Inicial de Vacaciones Fisicas no Valido ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFisVacacion(2).SetFocus: Exit Sub
  If CDec(txtVacaciones(0).Text) = 0 And mskFisVacacion(3).ClipText <> "" Then Beep: MsgBox "Fecha Final de Vacaciones Fisicas no Valido ", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFisVacacion(3).SetFocus: Exit Sub
  ' Vacaciones Fisicas - periodos
  If mskFisVacacion(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFisVacacion(0), 1900) Then tasRegister.Tabs(1).Selected = True: mskFisVacacion(0).SetFocus: Exit Sub
  End If
  If mskFisVacacion(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFisVacacion(1), 1900) Then tasRegister.Tabs(1).Selected = True: mskFisVacacion(1).SetFocus: Exit Sub
  End If
  If mskFisVacacion(2).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFisVacacion(2), 1900) Then tasRegister.Tabs(1).Selected = True: mskFisVacacion(2).SetFocus: Exit Sub
  End If
  If mskFisVacacion(3).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFisVacacion(3), 1900) Then tasRegister.Tabs(1).Selected = True: mskFisVacacion(3).SetFocus: Exit Sub
  End If
  
  ' Venta Vacaciones  - información
  If CDec(txtVacaciones(1).Text) > 0 And Len(mskPerVacacion(2).ClipText) < 8 Then Beep: MsgBox "Debe Ingresar Periodo de Venta Vacaciones", vbExclamation: tasRegister.Tabs(1).Selected = True: mskPerVacacion(2).SetFocus: Exit Sub
  If CDec(txtVacaciones(1).Text) = 0 And Len(mskPerVacacion(2).ClipText) <> 0 Then Beep: MsgBox "Periodo de Venta Vacaciones no Valido", vbExclamation: tasRegister.Tabs(1).Selected = True: mskPerVacacion(2).SetFocus: Exit Sub
  If CDec(txtVacaciones(1).Text) > 0 And mskFisVacacion(4).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicial Venta Vacaciones", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFisVacacion(4).SetFocus: Exit Sub
  If CDec(txtVacaciones(1).Text) > 0 And mskFisVacacion(5).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final Venta Vacaciones", vbExclamation: tasRegister.Tabs(1).Selected = True: mskFisVacacion(5).SetFocus: Exit Sub
  ' Venta Vacaciones  - periodos
  If mskFisVacacion(4).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFisVacacion(4), 1900) Then tasRegister.Tabs(1).Selected = True: mskFisVacacion(4).SetFocus: Exit Sub
  End If
  If mskFisVacacion(5).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFisVacacion(5), 1900) Then tasRegister.Tabs(1).Selected = True: mskFisVacacion(5).SetFocus: Exit Sub
  End If
  
  ' PRIMERA FICHA - Licencias
  ' Pre y Post Natal - información y fechas
  If CDec(txtPrePostNatal.Text) > 0 And txtMotivo(1).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(1).SetFocus: Exit Sub
  If CDec(txtPrePostNatal.Text) > 0 And lblHelp(1).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(1).SetFocus: Exit Sub
  If CDec(txtPrePostNatal.Text) > 0 And mskFecPreNatal(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Pre-Post Natal", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecPreNatal(0).SetFocus: Exit Sub
  If CDec(txtPrePostNatal.Text) > 0 And mskFecPreNatal(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Pre-Post Natal", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecPreNatal(1).SetFocus: Exit Sub
  If CDec(txtPrePostNatal.Text) > 0 And txtCertificado(0).Text = "" Then Beep: MsgBox "Debe Ingresar Certificado de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(2).Selected = True: txtCertificado(0).SetFocus: Exit Sub
  If mskFecPreNatal(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecPreNatal(0), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecPreNatal(0).SetFocus: Exit Sub
  End If
  If mskFecPreNatal(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecPreNatal(1), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecPreNatal(1).SetFocus: Exit Sub
  End If
  
  ' Enfermedad Subsidio - información y fechas
  If CDec(txtEnfermedad.Text) > 0 And txtMotivo(3).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(3).SetFocus: Exit Sub
  If CDec(txtEnfermedad.Text) > 0 And lblHelp(3).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(3).SetFocus: Exit Sub
  If CDec(txtEnfermedad.Text) > 0 And mskFecEnfermedad(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Enfermedad", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecEnfermedad(0).SetFocus: Exit Sub
  If CDec(txtEnfermedad.Text) > 0 And mskFecEnfermedad(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Enfermedad", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecEnfermedad(1).SetFocus: Exit Sub
  If CDec(txtEnfermedad.Text) > 0 And txtCertificado(1).Text = "" Then Beep: MsgBox "Debe Ingresar Certificado de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(2).Selected = True: txtCertificado(1).SetFocus: Exit Sub
  If mskFecEnfermedad(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecEnfermedad(0), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecEnfermedad(0).SetFocus: Exit Sub
  End If
  If mskFecEnfermedad(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecEnfermedad(1), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecEnfermedad(1).SetFocus: Exit Sub
  End If
  
  ' Accidente Desanso medico - información y fechas
  If CDec(txtAccidente.Text) > 0 And txtMotivo(2).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(2).SetFocus: Exit Sub
  If CDec(txtAccidente.Text) > 0 And lblHelp(2).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(2).SetFocus: Exit Sub
  If CDec(txtAccidente.Text) > 0 And mskFecAccidente(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Accidente", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecAccidente(0).SetFocus: Exit Sub
  If CDec(txtAccidente.Text) > 0 And mskFecAccidente(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Accidente", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecAccidente(1).SetFocus: Exit Sub
  If mskFecAccidente(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecAccidente(0), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecAccidente(0).SetFocus: Exit Sub
  End If
  If mskFecAccidente(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecAccidente(1), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecAccidente(1).SetFocus: Exit Sub
  End If
  
  ' Licencia sin goce - información y fechas
  If CDec(txtLicencia.Text) > 0 And txtMotivo(4).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(4).SetFocus: Exit Sub
  If CDec(txtLicencia.Text) > 0 And lblHelp(4).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(2).Selected = True: txtMotivo(4).SetFocus: Exit Sub
  If CDec(txtLicencia.Text) > 0 And mskFecLicencia(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Licencia", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecLicencia(0).SetFocus: Exit Sub
  If CDec(txtLicencia.Text) > 0 And mskFecLicencia(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Licencia", vbExclamation: tasRegister.Tabs(2).Selected = True: mskFecLicencia(1).SetFocus: Exit Sub
  If mskFecLicencia(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecLicencia(0), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecLicencia(0).SetFocus: Exit Sub
  End If
  If mskFecLicencia(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecLicencia(1), 1900) Then tasRegister.Tabs(2).Selected = True: mskFecLicencia(1).SetFocus: Exit Sub
  End If

  ' SEGUNDA FICHA - Liquidación
  ' Liquidación - información
  If mskFechaCese.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFechaCese, 1900) Then tasRegister.Tabs(3).Selected = True: mskFechaCese.SetFocus: Exit Sub
  End If
  ' Vacaciones liquidaciones - fechas
  If mskFecVacaLiqui(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacaLiqui(0), 1900) Then tasRegister.Tabs(3).Selected = True: mskFecVacaLiqui(0).SetFocus: Exit Sub
  End If
  If mskFecVacaLiqui(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacaLiqui(1), 1900) Then tasRegister.Tabs(3).Selected = True: mskFecVacaLiqui(1).SetFocus: Exit Sub
  End If
  
  ' Vacaciones Vencidas - información
  If CDec(txtVacaVencida.Text) > 0 And Len(mskPerVacacion(3).ClipText) < 8 Then Beep: MsgBox "Debe Ingresar Periodo Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskPerVacacion(3).SetFocus: Exit Sub
  If CDec(txtVacaVencida.Text) = 0 And Len(mskPerVacacion(3).ClipText) <> 0 Then Beep: MsgBox "Periodo Vacaciones Vencidas no Valido", vbExclamation: tasRegister.Tabs(3).Selected = True: mskPerVacacion(3).SetFocus: Exit Sub
  If CDec(txtVacaVencida.Text) > 0 And mskFecVacaVen(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicial Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskFecVacaVen(0).SetFocus: Exit Sub
  If CDec(txtVacaVencida.Text) > 0 And mskFecVacaVen(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskFecVacaVen(1).SetFocus: Exit Sub
  
  If CDec(txtVacaVencida.Text) = 0 And Len(mskPerVacacion(4).ClipText) <> 0 Then Beep: MsgBox "Periodo Vacaciones Vencidas no Valido", vbExclamation: tasRegister.Tabs(3).Selected = True: mskPerVacacion(4).SetFocus: Exit Sub
  If CDec(txtVacaVencida.Text) > 0 And (Len(mskPerVacacion(4).ClipText) <> 0 And Len(mskPerVacacion(4).ClipText) < 8) Then Beep: MsgBox "Periodo Vacaciones Vencidas no Valido", vbExclamation: tasRegister.Tabs(3).Selected = True: mskPerVacacion(4).SetFocus: Exit Sub
  If Len(mskPerVacacion(4).ClipText) <> 0 And mskFecVacaVen(2).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicial Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskFecVacaVen(2).SetFocus: Exit Sub
  If Len(mskPerVacacion(4).ClipText) <> 0 And mskFecVacaVen(3).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskFecVacaVen(3).SetFocus: Exit Sub
  If Len(mskPerVacacion(4).ClipText) = 0 And mskFecVacaVen(2).ClipText <> "" Then Beep: MsgBox "Debe Ingresar Periodo Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskPerVacacion(4).SetFocus: Exit Sub
  If Len(mskPerVacacion(4).ClipText) = 0 And mskFecVacaVen(3).ClipText <> "" Then Beep: MsgBox "Debe Ingresar Periodo Vacaciones Vencidas", vbExclamation: tasRegister.Tabs(3).Selected = True: mskPerVacacion(4).SetFocus: Exit Sub
  ' Vacaciones Vencidas - periodos
  If mskFecVacaVen(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacaVen(0), 1900) Then tasRegister.Tabs(3).Selected = True: mskFecVacaVen(0).SetFocus: Exit Sub
  End If
  If mskFecVacaVen(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacaVen(1), 1900) Then tasRegister.Tabs(3).Selected = True: mskFecVacaVen(1).SetFocus: Exit Sub
  End If
  If mskFecVacaVen(2).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacaVen(2), 1900) Then tasRegister.Tabs(3).Selected = True: mskFecVacaVen(2).SetFocus: Exit Sub
  End If
  If mskFecVacaVen(3).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecVacaVen(3), 1900) Then tasRegister.Tabs(3).Selected = True: mskFecVacaVen(3).SetFocus: Exit Sub
  End If
  
  ' TERCERA FICHA - Permisos
  ' Licencia Paternidad - información y fechas
  If CDec(txtpaternidad.Text) > 0 And txtMotivo(5).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(4).Selected = True: txtMotivo(5).SetFocus: Exit Sub
  If CDec(txtpaternidad.Text) > 0 And lblHelp(5).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(4).Selected = True: txtMotivo(5).SetFocus: Exit Sub
  If CDec(txtpaternidad.Text) > 0 And mskFecPaternidad(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Paternidad", vbExclamation: tasRegister.Tabs(4).Selected = True: mskFecPaternidad(0).SetFocus: Exit Sub
  If CDec(txtpaternidad.Text) > 0 And mskFecPaternidad(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Paternidad", vbExclamation: tasRegister.Tabs(4).Selected = True: mskFecPaternidad(1).SetFocus: Exit Sub
  If mskFecPaternidad(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecPaternidad(0), 1900) Then tasRegister.Tabs(4).Selected = True: mskFecPaternidad(0).SetFocus: Exit Sub
  End If
  If mskFecPaternidad(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecPaternidad(1), 1900) Then tasRegister.Tabs(4).Selected = True: mskFecPaternidad(1).SetFocus: Exit Sub
  End If
  
  ' Licencia Fallecimiento - información y fechas
  If CDec(txtFallece.Text) > 0 And txtMotivo(6).Text = "" Then Beep: MsgBox "Debe Ingresar Motivo de Inasistencia Laboral", vbExclamation: tasRegister.Tabs(4).Selected = True: txtMotivo(6).SetFocus: Exit Sub
  If CDec(txtFallece.Text) > 0 And lblHelp(6).Caption = "???" Then Beep: MsgBox "Motivo de Inasistencia Laboral no Valido; Verificar", vbExclamation: tasRegister.Tabs(4).Selected = True: txtMotivo(6).SetFocus: Exit Sub
  If CDec(txtFallece.Text) > 0 And mskFecFallece(0).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Inicio de Fallecimiento Familiar", vbExclamation: tasRegister.Tabs(4).Selected = True: mskFecFallece(0).SetFocus: Exit Sub
  If CDec(txtFallece.Text) > 0 And mskFecFallece(1).ClipText = "" Then Beep: MsgBox "Debe Ingresar la Fecha Final de Fallecimiento Familiar", vbExclamation: tasRegister.Tabs(4).Selected = True: mskFecFallece(1).SetFocus: Exit Sub
  If mskFecFallece(0).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecFallece(0), 1900) Then tasRegister.Tabs(4).Selected = True: mskFecFallece(0).SetFocus: Exit Sub
  End If
  If mskFecFallece(1).ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecFallece(1), 1900) Then tasRegister.Tabs(4).Selected = True: mskFecFallece(1).SetFocus: Exit Sub
  End If
  
  ' Validar dias de Vacaciones no sea mayor a 30 dias por periodo
  ' Vacaciones fisicas periodo 1 y 2, Venta de vacaciones, Vacaciones vencidas periodo 1 y 2
  s_Sql = ""
  For nSecuencia = 1 To 5
    s_Sql = s_Sql & "SELECT asi.codpsn, asi.pdovaca" & nSecuencia & " AS Periodo_Vacaciones, "
    s_Sql = s_Sql & "IFNULL(SUM(DATEDIFF(asi.fechafinvaca" & nSecuencia & ",asi.fechainivaca" & nSecuencia & ")+1),0) AS Dias_Vacaciones "
    s_Sql = s_Sql & "FROM plasistencia asi "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo per ON asi.codcls=per.codcls AND asi.codpdo=per.codpdo "
    s_Sql = s_Sql & "INNER JOIN plperiodo qry ON per.codcls=qry.codcls AND qry.codpdo='" & Trim(o_SelAsistencia.txtPeriodo.Text) & "' AND CONCAT(per.anopdo,per.mespdo)<CONCAT(qry.anopdo,qry.mespdo) "
    s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND asi.codpsn='" & Trim(o_SelAsistencia.dcaRegistro.Recordset!codpsn) & "' "
    s_Sql = s_Sql & "AND asi.fechainivaca" & nSecuencia & ">=psn.fecingreso "
    s_Sql = s_Sql & "AND IFNULL(asi.pdovaca" & nSecuencia & ",'')<>'' "
    s_Sql = s_Sql & "AND asi.pdovaca" & nSecuencia & " IN('" & Trim(mskPerVacacion(0).ClipText) & "', '" & Trim(mskPerVacacion(1).ClipText) & "', '" & Trim(mskPerVacacion(2).ClipText) & "', '" & Trim(mskPerVacacion(3).ClipText) & "', '" & Trim(mskPerVacacion(4).ClipText) & "') "
    s_Sql = s_Sql & "GROUP BY asi.codpsn, asi.pdovaca" & nSecuencia & " "
    s_Sql = s_Sql & IIf(nSecuencia = 5, "", "UNION ")
  Next nSecuencia
  s_Sql = s_Sql & "ORDER BY codpsn, Periodo_Vacaciones"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Saldo anterior vacaciones fisicas
  nSecuencia = 0
  ReDim a_Vacaciones(2, nSecuencia)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    While Not porstRecordset.EOF
      If nSecuencia > 0 Then
        nRegistroID = fRetornaPosArreglo(a_Vacaciones, 1, 1, Trim(porstRecordset("Periodo_Vacaciones")))
      End If
      ' Redimensiono arreglo periodo vacacional
      If nRegistroID = 0 Then
        nSecuencia = nSecuencia + 1
        ReDim Preserve a_Vacaciones(2, nSecuencia)
      End If
      nRegistroID = nSecuencia
      a_Vacaciones(1, nRegistroID) = Trim(porstRecordset("Periodo_Vacaciones"))
      a_Vacaciones(2, nRegistroID) = a_Vacaciones(2, nRegistroID) + CDec(porstRecordset("Dias_Vacaciones"))
      porstRecordset.MoveNext
    Wend
  End If
  porstRecordset.Close
   
  ' Valido limite de dias por periodo
  nSecuencia = 0
  ReDim a_AcumVacacion(2, nSecuencia)
  
  ' Vacaciones fisicas y venta de vacaciones
  nIndexIni = 0: nIndexFin = -1
  For nRegistroIX = 0 To 2
    nIndexIni = nIndexFin + 1
    nIndexFin = nIndexIni + 1
    If IsDate(mskFisVacacion(nIndexIni).FormattedText) And IsDate(mskFisVacacion(nIndexFin).FormattedText) Then
      nDiaVacacion = IIf(IsNull(DateDiff("d", mskFisVacacion(nIndexIni).FormattedText, mskFisVacacion(nIndexFin).FormattedText)), 0, DateDiff("d", mskFisVacacion(nIndexIni).FormattedText, mskFisVacacion(nIndexFin).FormattedText) + 1)
      nDAcuVacacion = 0: nRegistroID = 0
      If UBound(a_Vacaciones, 2) > 0 Then
        nRegistroID = fRetornaPosArreglo(a_Vacaciones, 1, 1, Trim(mskPerVacacion(nRegistroIX).ClipText))
        If nRegistroID > 0 Then
          nDAcuVacacion = a_Vacaciones(2, nRegistroID)
        End If
      End If
      
      ' Redimensiono arreglo acumulado vacaciones
      nRegistroID = fRetornaPosArreglo(a_AcumVacacion, 1, 1, Trim(mskPerVacacion(nRegistroIX).ClipText))
      If nRegistroID = 0 Then
        nSecuencia = nSecuencia + 1
        ReDim Preserve a_AcumVacacion(2, nSecuencia)
      End If
      nRegistroID = nSecuencia
      a_AcumVacacion(1, nRegistroID) = Trim(mskPerVacacion(nRegistroIX).ClipText)
      a_AcumVacacion(2, nRegistroID) = a_AcumVacacion(2, nRegistroID) + CDec(nDiaVacacion)
      
     ' Valido periódo supere los 30 días
     If CDec(a_AcumVacacion(2, nRegistroID)) > 30 Then
       MsgBox "El numero de días de Vacaciones acumuladas :" & Chr(13) & "Periodo : " & "'" & Trim(mskPerVacacion(nRegistroIX).ClipText) & "'" & " es de :" & "'" & nDAcuVacacion & "'" & " días" & Chr(13) & "Los días que se esta asignando para este periódo son de : " & "'" & a_AcumVacacion(2, nRegistroID) & "'" & Chr(13) & "Datos ingresados no validos, no se puede Grabar", vbInformation
       Exit Sub
     End If
    End If
  Next nRegistroIX
  
  ' Vacaciones vencidas
  nIndexIni = 0: nIndexFin = -1
  For nRegistroIX = 3 To 4
    nIndexIni = nIndexFin + 1
    nIndexFin = nIndexIni + 1
    If IsDate(mskFecVacaVen(nIndexIni).FormattedText) And IsDate(mskFecVacaVen(nIndexFin).FormattedText) Then
      nDiaVacacion = IIf(IsNull(DateDiff("d", mskFecVacaVen(nIndexIni).FormattedText, mskFecVacaVen(nIndexFin).FormattedText)), 0, DateDiff("d", mskFecVacaVen(nIndexIni).FormattedText, mskFecVacaVen(nIndexFin).FormattedText) + 1)
      nDAcuVacacion = 0: nRegistroID = 0
      If UBound(a_Vacaciones, 2) > 0 Then
        nRegistroID = fRetornaPosArreglo(a_Vacaciones, 1, 1, Trim(mskPerVacacion(nRegistroIX).ClipText))
        If nRegistroID > 0 Then
          nDAcuVacacion = a_Vacaciones(2, nRegistroID)
        End If
      End If
      
      ' Redimensiono arreglo acumulado vacaciones
      nRegistroID = fRetornaPosArreglo(a_AcumVacacion, 1, 1, Trim(mskPerVacacion(nRegistroIX).ClipText))
      If nRegistroID = 0 Then
        nSecuencia = nSecuencia + 1
        ReDim Preserve a_AcumVacacion(2, nSecuencia)
      End If
      nRegistroID = nSecuencia
      a_AcumVacacion(1, nRegistroID) = Trim(mskPerVacacion(nRegistroIX).ClipText)
      a_AcumVacacion(2, nRegistroID) = a_AcumVacacion(2, nRegistroID) + CDec(nDiaVacacion)
      
     ' Valido periódo supere los 30 días
     If CDec(a_AcumVacacion(2, nRegistroID)) > 30 Then
       MsgBox "El numero de días de Vacaciones acumuladas :" & Chr(13) & "Periodo : " & "'" & Trim(mskPerVacacion(nRegistroIX).ClipText) & "'" & " es de :" & "'" & nDAcuVacacion & "'" & " días" & Chr(13) & "Los días que se esta asignando para este periódo son de : " & "'" & a_AcumVacacion(2, nRegistroID) & "'" & Chr(13) & "Datos ingresados no validos, no se puede Grabar", vbInformation
       Exit Sub
     End If
    End If
  Next nRegistroIX
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Capturo el registro a actualizar
  s_Registro = Trim(o_SelAsistencia.dcaRegistro.Recordset!codpsn)
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpdo", "codpsn", "diatrabajo", "diamediotm", "diaparcial", "dialaboral", "diafalta", "tardanza", "diaferiado", "diatradesemanal", "diasuspension", "dialibre", "horanormal", "horamediotm", "horaparcial", "horatipo1", "horatipo2", "horatipo3", "horatipo4", "opcional", _
                  "indvacadelanta", "diavacaciones", "codmdi_vacac", "fechainivacacion", "fechafinvacacion", "pdovaca1", "fechainivaca1", "fechafinvaca1", "pdovaca2", "fechainivaca2", "fechafinvaca2", "diavacaventa", "pdovaca3", "fechainivaca3", "fechafinvaca3", _
                  "diaprepostnatal", "codmdi_natal", "fechaini_natal", "fechafin_natal", "numecitt_natal", "accidente", "codmdi_accid", "fechaini_accid", "fechafin_accid", "enfermedad", "codmdi_enfer", "fechaini_enfer", "fechafin_enfer", "numecitt_enfer", "licencia", "codmdi_licen", "fechaini_licen", "fechafin_licen", _
                  "fechacese", "dialiquidacion", "diagratificacion", "liquidavacacion", "fechainiliqvaca", "fechafinliqvaca", "observacion", "liqnocalifica", "diavacavencida", "pdovaca4", "fechainivaca4", "fechafinvaca4", "pdovaca5", "fechainivaca5", "fechafinvaca5", _
                  "diapaternidad", "codmdi_pater", "fechaini_pater", "fechafin_pater", "diafallecefam", "codmdi_falle", "fechaini_falle", "fechafin_falle", _
                  IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, Trim(o_SelAsistencia.txtPeriodo.Text), Trim(o_SelAsistencia.dcaRegistro.Recordset!codpsn), CDec(txtDiaTrabajo.Text), CDec(txtDiaMedio.Text), CDec(txtDiaParcial.Text), CDec(txtDiaLaboral.Text), CDec(txtFaltas.Text), CDec(txtTardanzas.Text), CDec(txtDiaFeriado.Text), CDec(txtDiaSemanal.Text), CDec(txtDiaSuspension.Text), CDec(txtDiaLibre.Text), CDec(txtHoraNormal.Text), CDec(txtHoraMedio.Text), CDec(txtHoraParcial.Text), CDec(txtHrExtraSimple.Text), CDec(txtHrExtraDoble.Text), CDec(txtHrEspecial.Text), CDec(txtHrNocturno.Text), CDec(txtOpcional.Text), _
                  IIf(chkAdeVacacion.Value, s_Estado_Act, s_Estado_Ina), CDec(txtVacaciones(0).Text), Trim(txtMotivo(0).Text), Format(mskFecVacacion(0), s_FmtFechMysql_0), Format(mskFecVacacion(1), s_FmtFechMysql_0), Trim(mskPerVacacion(0).ClipText), Format(mskFisVacacion(0), s_FmtFechMysql_0), Format(mskFisVacacion(1), s_FmtFechMysql_0), Trim(mskPerVacacion(1).ClipText), Format(mskFisVacacion(2), s_FmtFechMysql_0), Format(mskFisVacacion(3), s_FmtFechMysql_0), CDec(txtVacaciones(1).Text), Trim(mskPerVacacion(2).ClipText), Format(mskFisVacacion(4), s_FmtFechMysql_0), Format(mskFisVacacion(5), s_FmtFechMysql_0), _
                  CDec(txtPrePostNatal.Text), Trim(txtMotivo(1).Text), Format(mskFecPreNatal(0), s_FmtFechMysql_0), Format(mskFecPreNatal(1), s_FmtFechMysql_0), Trim(txtCertificado(0).Text), _
                  CDec(txtAccidente.Text), Trim(txtMotivo(2).Text), Format(mskFecAccidente(0), s_FmtFechMysql_0), Format(mskFecAccidente(1), s_FmtFechMysql_0), _
                  CDec(txtEnfermedad.Text), Trim(txtMotivo(3).Text), Format(mskFecEnfermedad(0), s_FmtFechMysql_0), Format(mskFecEnfermedad(1), s_FmtFechMysql_0), Trim(txtCertificado(1).Text), _
                  CDec(txtLicencia.Text), Trim(txtMotivo(4).Text), Format(mskFecLicencia(0), s_FmtFechMysql_0), Format(mskFecLicencia(1), s_FmtFechMysql_0), _
                  Format(mskFechaCese, s_FmtFechMysql_0), CDec(txtDiasLiquidacion.Text), CDec(txtDiasLiqGratifica.Text), CDec(txtDiasLiqVacacion.Text), Format(mskFecVacaLiqui(0), s_FmtFechMysql_0), Format(mskFecVacaLiqui(1), s_FmtFechMysql_0), Trim(txtObservacion.Text), IIf(chkLiquidaCalifica.Value, s_Estado_Act, s_Estado_Ina), _
                  CDec(txtVacaVencida.Text), Trim(mskPerVacacion(3).ClipText), Format(mskFecVacaVen(0), s_FmtFechMysql_0), Format(mskFecVacaVen(1), s_FmtFechMysql_0), Trim(mskPerVacacion(4).ClipText), Format(mskFecVacaVen(2), s_FmtFechMysql_0), Format(mskFecVacaVen(3), s_FmtFechMysql_0), _
                  CDec(txtpaternidad.Text), Trim(txtMotivo(5).Text), Format(mskFecPaternidad(0), s_FmtFechMysql_0), Format(mskFecPaternidad(1), s_FmtFechMysql_0), CDec(txtFallece.Text), Trim(txtMotivo(6).Text), Format(mskFecFallece(0), s_FmtFechMysql_0), Format(mskFecFallece(1), s_FmtFechMysql_0), _
                  ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, _
                  TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, _
                  TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, _
                  TipoDato.FECHA, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, _
                  TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, _
                  TipoDato.Caracter, TipoDato.FECHA)
  a_Where = Array("codcls", "codpdo", "codpsn")
   

  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Procesando Información ..."
  ' Incializo rango de actualización
  If l_Inicializa Then
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    Set porstClone = o_SelAsistencia.dcaRegistro.Recordset.Clone()
    For n_Index = 0 To o_SelAsistencia.tdbRegistro.SelBookmarks.Count - 1
      porstClone.Bookmark = o_SelAsistencia.tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, "iniasispsn", porstClone!codpsn, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    ' Cierro, elimino de memoria objeto
    porstClone.Close
    Set porstClone = Nothing
  End If
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If l_Inicializa Then
      ' Elimino la asistencia del rango
      s_Sql = "DELETE asi "
      s_Sql = s_Sql & "FROM plasistencia asi, rangoimpresion rng "
      s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND asi.codpdo='" & Trim(o_SelAsistencia.txtPeriodo.Text) & "' "
      s_Sql = s_Sql & "AND asi.codpsn=rng.valor "
      s_Sql = s_Sql & "AND rng.proceso='iniasispsn' "
      s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND rng.fyhcre='" & s_FechaHora & "'"
      If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
      ' Actualizo la asistencia del rango
      s_Sql = "INSERT INTO plasistencia ("
      For n_Index = 0 To UBound(a_Campos)
        s_Sql = s_Sql & a_Campos(n_Index)
        s_Sql = s_Sql & IIf(n_Index <> UBound(a_Campos), ", ", ") ")
      Next n_Index
      
      s_Sql = s_Sql & "SELECT '" & ps_ClsPlanilla & "', '" & Trim(o_SelAsistencia.txtPeriodo.Text) & "', "
      s_Sql = s_Sql & "LEFT(rng.valor, 11), " & CDec(txtDiaTrabajo.Text) & ", " & CDec(txtDiaMedio.Text) & ", " & CDec(txtDiaParcial.Text) & ", "
      s_Sql = s_Sql & CDec(txtDiaLaboral.Text) & ", " & CDec(txtFaltas.Text) & ", " & CDec(txtTardanzas.Text) & ", " & CDec(txtDiaFeriado.Text) & ", "
      s_Sql = s_Sql & CDec(txtDiaSemanal.Text) & ", " & CDec(txtDiaSuspension.Text) & ", " & CDec(txtDiaLibre.Text) & ", "
      s_Sql = s_Sql & "(CASE WHEN psn.jornadalaboral>0 THEN ROUND(" & CDec(txtDiaTrabajo.Text) & "*psn.jornadalaboral, 2) ELSE " & CDec(txtHoraNormal.Text) & " END), "
      s_Sql = s_Sql & CDec(txtHoraMedio.Text) & ", " & CDec(txtHoraParcial.Text) & ", " & CDec(txtHrExtraSimple.Text) & ", "
      s_Sql = s_Sql & CDec(txtHrExtraDoble.Text) & ", " & CDec(txtHrEspecial.Text) & ", " & CDec(txtHrNocturno.Text) & ", " & CDec(txtOpcional.Text) & ", "
      s_Sql = s_Sql & IIf(chkAdeVacacion.Value, s_Estado_Act, s_Estado_Ina) & ", " & CDec(txtVacaciones(0).Text) & ", " & IIf(Trim(txtMotivo(0).Text) = "", "Null", "'" & Trim(txtMotivo(0).Text) & "'") & ", "
      If IsDate(mskFecVacacion(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacacion(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecVacacion(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacacion(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & IIf(Trim(mskPerVacacion(0).ClipText) = "", "Null", "'" & Trim(mskPerVacacion(0).ClipText) & "'") & ", "
      If IsDate(mskFisVacacion(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFisVacacion(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFisVacacion(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFisVacacion(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & IIf(Trim(mskPerVacacion(1).ClipText) = "", "Null", "'" & Trim(mskPerVacacion(1).ClipText) & "'") & ", "
      If IsDate(mskFisVacacion(2)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFisVacacion(2), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFisVacacion(3)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFisVacacion(3), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & CDec(txtVacaciones(1).Text) & ", "
      s_Sql = s_Sql & IIf(Trim(mskPerVacacion(2).ClipText) = "", "Null", "'" & Trim(mskPerVacacion(2).ClipText) & "'") & ", "
      If IsDate(mskFisVacacion(4)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFisVacacion(4), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFisVacacion(5)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFisVacacion(5), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      
      ' Primera ficha
      s_Sql = s_Sql & CDec(txtPrePostNatal.Text) & ", " & IIf(Trim(txtMotivo(1).Text) = "", "Null", "'" & Trim(txtMotivo(1).Text) & "'") & ", "
      If IsDate(mskFecPreNatal(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecPreNatal(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecPreNatal(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecPreNatal(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & IIf(Trim(txtCertificado(0).Text) = "", "Null", "'" & Trim(txtCertificado(0).Text) & "'") & ", "
      
      s_Sql = s_Sql & CDec(txtAccidente.Text) & ", " & IIf(Trim(txtMotivo(2).Text) = "", "Null", "'" & Trim(txtMotivo(2).Text) & "'") & ", "
      If IsDate(mskFecAccidente(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecAccidente(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecAccidente(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecAccidente(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      
      s_Sql = s_Sql & CDec(txtEnfermedad.Text) & ", " & IIf(Trim(txtMotivo(3).Text) = "", "Null", "'" & Trim(txtMotivo(3).Text) & "'") & ", "
      If IsDate(mskFecEnfermedad(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecEnfermedad(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecEnfermedad(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecEnfermedad(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & IIf(Trim(txtCertificado(1).Text) = "", "Null", "'" & Trim(txtCertificado(1).Text) & "'") & ", "
      
      s_Sql = s_Sql & CDec(txtLicencia.Text) & ", " & IIf(Trim(txtMotivo(4).Text) = "", "Null", "'" & Trim(txtMotivo(4).Text) & "'") & ", "
      If IsDate(mskFecLicencia(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecLicencia(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecLicencia(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecLicencia(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      
      ' Segunda ficha
      If IsDate(mskFechaCese) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFechaCese, s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & CDec(txtDiasLiquidacion.Text) & ", " & CDec(txtDiasLiqGratifica.Text) & ", " & CDec(txtDiasLiqVacacion.Text) & ", "
      If IsDate(mskFecVacaLiqui(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacaLiqui(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecVacaLiqui(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacaLiqui(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & IIf(Trim(txtObservacion.Text) = "", "Null", "'" & txtObservacion.Text & "'") & ", "
      s_Sql = s_Sql & CInt(IIf(chkLiquidaCalifica.Value, s_Estado_Act, s_Estado_Ina)) & ", "
    
      s_Sql = s_Sql & CDec(txtVacaVencida.Text) & ", " & IIf(Trim(mskPerVacacion(3).ClipText) = "", "Null", "'" & Trim(mskPerVacacion(3).ClipText) & "'") & ", "
       If IsDate(mskFecVacaVen(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacaVen(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecVacaVen(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacaVen(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & IIf(Trim(mskPerVacacion(4).ClipText) = "", "Null", "'" & Trim(mskPerVacacion(4).ClipText) & "'") & ", "
       If IsDate(mskFecVacaVen(2)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacaVen(2), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecVacaVen(3)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecVacaVen(3), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
  
      ' Tercera ficha
      s_Sql = s_Sql & CDec(txtpaternidad.Text) & ", " & IIf(Trim(txtMotivo(5).Text) = "", "Null", "'" & Trim(txtMotivo(5).Text) & "'") & ", "
      If IsDate(mskFecPaternidad(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecPaternidad(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecPaternidad(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecPaternidad(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      s_Sql = s_Sql & CDec(txtFallece.Text) & ", " & IIf(Trim(txtMotivo(6).Text) = "", "Null", "'" & Trim(txtMotivo(6).Text) & "'") & ", "
      If IsDate(mskFecFallece(0)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecFallece(0), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
      If IsDate(mskFecFallece(1)) Then
        s_Sql = s_Sql & "DATE_FORMAT('" & Format(mskFecFallece(1), s_FmtFechMysql_0) & "', '" & s_FmtFechMysql_1 & "'), "
      Else
        s_Sql = s_Sql & "Null, "
      End If
   
      ' Sección final
      s_Sql = s_Sql & "'" & ps_Usuario & "', '" & Format(Now, s_FmtFeHoMysql_0) & "' "
      s_Sql = s_Sql & "FROM rangoimpresion rng "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls='" & ps_ClsPlanilla & "' AND LEFT(rng.valor, 11)=psn.codpsn "
      s_Sql = s_Sql & "WHERE rng.proceso='iniasispsn' "
      s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND rng.fyhcre='" & s_FechaHora & "' "
      s_Sql = s_Sql & "ORDER BY rng.valor"
      If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Error
      Me.Tag = s_MdoData_Vis
    Else
      If Not Records_Ins("plasistencia", a_Campos, a_Valores, a_Tipos) Then GoTo Error
    End If
  Else
    If Not Records_Upd("plasistencia", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle.Caption, vbInformation
  ' Ubico en el siguiente registro a ingresar o actualizar
  If Not l_Inicializa Then cmdMove_Click 2
  l_Inicializa = False
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtDiaTrabajo.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  MuestraMensaje s_OldMessage
  ' Elimino rango de inicialización
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, "iniasispsn", "", ps_Usuario, s_FechaHora, "E"
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 9540: Me.Width = 8640
  Me.Left = 2500: Me.Top = 1280
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Asistencia y Puntualidad"
  lblTitle = "Asistencia de Personal"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operación del registro
  l_Inicializa = (o_SelAsistencia.tdbRegistro.SelBookmarks.Count > 0)
  Me.Tag = IIf(l_Inicializa, s_MdoData_Ins, o_SelAsistencia.Tag)
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
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
  l_ExistRecord = (o_SelAsistencia.dcaRegistro.Recordset.EOF Or o_SelAsistencia.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = o_SelAsistencia.dcaRegistro.Recordset!codpsn
  
  ' Carga los datos en el formulario
  ShowScreen
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codtsu", "destsu")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Motivo de Inasistencia", 2
  ']
  
  ' Selecciono la primera pestaña
  n_IndexTabs = 2
  tasRegister.Tabs(1).Selected = True
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub mskFecAccidente_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecAccidente(Index)
End Sub
Private Sub mskFecAccidente_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecAccidente_Validate(Index As Integer, Cancel As Boolean)
  If mskFecAccidente(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecAccidente(Index), 1900
  End If
End Sub
Private Sub mskFecEnfermedad_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecEnfermedad(Index)
End Sub
Private Sub mskFecEnfermedad_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecEnfermedad_Validate(Index As Integer, Cancel As Boolean)
  If mskFecEnfermedad(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecEnfermedad(Index), 1900
  End If
End Sub
Private Sub mskFecFallece_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecFallece(Index)
End Sub
Private Sub mskFecFallece_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecFallece_Validate(Index As Integer, Cancel As Boolean)
  If mskFecFallece(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecFallece(Index), 1900
  End If
End Sub
Private Sub mskFechaCese_GotFocus()
  gdl_Procedure.MarcaGet mskFechaCese
End Sub
Private Sub mskFechaCese_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFechaCese_Validate(Cancel As Boolean)
  If mskFechaCese.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFechaCese, 1900
  End If
End Sub
Private Sub mskFecLicencia_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecLicencia(Index)
End Sub
Private Sub mskFecLicencia_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecLicencia_Validate(Index As Integer, Cancel As Boolean)
  If mskFecLicencia(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecLicencia(Index), 1900
  End If
End Sub
Private Sub mskFecPaternidad_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecPaternidad(Index)
End Sub
Private Sub mskFecPaternidad_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecPaternidad_Validate(Index As Integer, Cancel As Boolean)
  If mskFecPaternidad(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecPaternidad(Index), 1900
  End If
End Sub
Private Sub mskFecPreNatal_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecPreNatal(Index)
End Sub
Private Sub mskFecPreNatal_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecPreNatal_Validate(Index As Integer, Cancel As Boolean)
  If mskFecPreNatal(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecPreNatal(Index), 1900
  End If
End Sub
Private Sub mskFecVacacion_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecVacacion(Index)
End Sub
Private Sub mskFecVacacion_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecVacacion_Validate(Index As Integer, Cancel As Boolean)
  If mskFecVacacion(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecVacacion(Index), 1900
  End If
End Sub
Private Sub mskFecVacaLiqui_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecVacaLiqui(Index)
End Sub
Private Sub mskFecVacaLiqui_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecVacaLiqui_Validate(Index As Integer, Cancel As Boolean)
  If mskFecVacaLiqui(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecVacaLiqui(Index), 1900
  End If
End Sub
Private Sub mskFecVacaVen_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFecVacaVen(Index)
End Sub
Private Sub mskFecVacaVen_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFecVacaVen_Validate(Index As Integer, Cancel As Boolean)
  If mskFecVacaVen(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecVacaVen(Index), 1900
  End If
End Sub
Private Sub mskFisVacacion_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskFisVacacion(Index)
End Sub
Private Sub mskFisVacacion_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub mskFisVacacion_Validate(Index As Integer, Cancel As Boolean)
  If mskFisVacacion(Index).ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFisVacacion(Index), 1900
  End If
End Sub
Private Sub mskPerVacacion_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet mskPerVacacion(Index)
End Sub
Private Sub mskPerVacacion_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
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
Private Sub tasRegister_Click()
  fraFicha(n_IndexTabs).Left = -20000
  fraFicha(n_IndexTabs).Enabled = False
  n_IndexTabs = (tasRegister.SelectedItem.Index - 1)
  fraFicha(n_IndexTabs).Top = 480
  fraFicha(n_IndexTabs).Left = 150
  fraFicha(n_IndexTabs).Enabled = True
End Sub
Private Sub tdbHelp_DblClick()
  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then Beep: MsgBox "No existen Registros para Seleccionar", vbExclamation: Exit Sub
  Select Case n_IndexHelp
   Case 0, 1, 2, 3, 4, 5, 6 ' motivo inasistencia
    txtMotivo(n_IndexHelp).Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = Left(tdbHelp.Columns(1).Value, 30)
    txtMotivo(n_IndexHelp).SetFocus
  End Select
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0, 1, 2, 3, 4, 5, 6 ' motivo inasistencia
    s_Sql = gdl_Funcion.HelpTablas("tsu", tdbHelp.Columns(ColIndex).DataField, "", "")
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
Private Sub txtAccidente_GotFocus()
  gdl_Procedure.MarcaGet txtAccidente
End Sub
Private Sub txtAccidente_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtAccidente_Validate(Cancel As Boolean)
  txtAccidente.Text = IIf(Not IsNumeric(txtAccidente.Text), 0, txtAccidente.Text)
  If CDec(txtAccidente.Text) < 0 Then MsgBox "Dias de Accidente no puede ser negativo; Verifique", vbInformation: txtAccidente.SetFocus: Exit Sub
  txtAccidente.Text = FormatNumber(txtAccidente.Text, 2)
End Sub
Private Sub txtCertificado_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCertificado(Index)
End Sub
Private Sub txtCertificado_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaFeriado_GotFocus()
  gdl_Procedure.MarcaGet txtDiaFeriado
End Sub
Private Sub txtDiaFeriado_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaFeriado_Validate(Cancel As Boolean)
  txtDiaFeriado.Text = IIf(Not IsNumeric(txtDiaFeriado.Text), 0, txtDiaFeriado.Text)
  If CDec(txtDiaFeriado.Text) < 0 Then MsgBox "Días Feriados no puede ser negativo; Verifique", vbInformation: txtDiaFeriado.SetFocus: Exit Sub
  txtDiaFeriado.Text = FormatNumber(txtDiaFeriado.Text, 2)
End Sub
Private Sub txtDiaLaboral_GotFocus()
  gdl_Procedure.MarcaGet txtDiaLaboral
End Sub
Private Sub txtDiaLaboral_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaLaboral_Validate(Cancel As Boolean)
  txtDiaLaboral.Text = IIf(Not IsNumeric(txtDiaLaboral.Text), 0, txtDiaLaboral.Text)
  If CDec(txtDiaLaboral.Text) < 0 Then MsgBox "Dias Laborables no puede ser negativo; Verifique", vbInformation: txtDiaLaboral.SetFocus: Exit Sub
  txtDiaLaboral.Text = FormatNumber(txtDiaLaboral.Text, 2)
  If (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd) And (txtDiaLaboral.Text <> txtDiaLaboral.Tag) Then txtHoraNormal.Text = FormatNumber(CDec(txtDiaLaboral.Text) * n_JornadaLaboral, 2)
  txtDiaLaboral.Tag = txtDiaLaboral.Text
End Sub
Private Sub txtDiaLibre_GotFocus()
  gdl_Procedure.MarcaGet txtDiaLibre
End Sub
Private Sub txtDiaLibre_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaLibre_Validate(Cancel As Boolean)
  txtDiaLibre.Text = IIf(Not IsNumeric(txtDiaLibre.Text), 0, txtDiaLibre.Text)
  If CDec(txtDiaLibre.Text) < 0 Then MsgBox "Dias libres no puede ser negativo; Verifique", vbInformation: txtDiaLibre.SetFocus: Exit Sub
  txtDiaLibre.Text = FormatNumber(txtDiaLibre.Text, 2)
End Sub
Private Sub txtDiaMedio_GotFocus()
  gdl_Procedure.MarcaGet txtDiaMedio
End Sub
Private Sub txtDiaMedio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaMedio_Validate(Cancel As Boolean)
  txtDiaMedio.Text = IIf(Not IsNumeric(txtDiaMedio.Text), 0, txtDiaMedio.Text)
  If CDec(txtDiaMedio.Text) < 0 Then MsgBox "Dias de Medio Tiempo no puede ser negativo; Verifique", vbInformation: txtDiaMedio.SetFocus: Exit Sub
  txtDiaMedio.Text = FormatNumber(txtDiaMedio.Text, 2)
End Sub
Private Sub txtDiaParcial_GotFocus()
  gdl_Procedure.MarcaGet txtDiaParcial
End Sub
Private Sub txtDiaParcial_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaParcial_Validate(Cancel As Boolean)
  txtDiaParcial.Text = IIf(Not IsNumeric(txtDiaParcial.Text), 0, txtDiaParcial.Text)
  If CDec(txtDiaParcial.Text) < 0 Then MsgBox "Dias de Tiempo Parcial no puede ser negativo; Verifique", vbInformation: txtDiaParcial.SetFocus: Exit Sub
  txtDiaParcial.Text = FormatNumber(txtDiaParcial.Text, 2)
End Sub
Private Sub txtDiaSemanal_GotFocus()
  gdl_Procedure.MarcaGet txtDiaSemanal
End Sub
Private Sub txtDiaSemanal_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaSemanal_Validate(Cancel As Boolean)
  txtDiaSemanal.Text = IIf(Not IsNumeric(txtDiaSemanal.Text), 0, txtDiaSemanal.Text)
  If CDec(txtDiaSemanal.Text) < 0 Then MsgBox "Dias de Trabajo Descanso Semanal Obligatorio no puede ser negativo; Verifique", vbInformation: txtDiaSemanal.SetFocus: Exit Sub
  txtDiaSemanal.Text = FormatNumber(txtDiaSemanal.Text, 2)
End Sub
Private Sub txtDiasLiqGratifica_GotFocus()
  gdl_Procedure.MarcaGet txtDiasLiqGratifica
End Sub
Private Sub txtDiasLiqGratifica_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiasLiqGratifica_Validate(Cancel As Boolean)
  txtDiasLiqGratifica.Text = IIf(Not IsNumeric(txtDiasLiqGratifica.Text), 0, txtDiasLiqGratifica.Text)
  If CDec(txtDiasLiqGratifica.Text) < 0 Then MsgBox "Dias de Liquidación de Gratificación no puede ser negativo; Verifique", vbInformation: txtDiasLiqGratifica.SetFocus: Exit Sub
  txtDiasLiqGratifica.Text = FormatNumber(txtDiasLiqGratifica.Text, 2)
End Sub
Private Sub txtDiasLiquidacion_GotFocus()
  gdl_Procedure.MarcaGet txtDiasLiquidacion
End Sub
Private Sub txtDiasLiquidacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiasLiquidacion_Validate(Cancel As Boolean)
  txtDiasLiquidacion.Text = IIf(Not IsNumeric(txtDiasLiquidacion.Text), 0, txtDiasLiquidacion.Text)
  If CDec(txtDiasLiquidacion.Text) < 0 Then MsgBox "Dias de Liquidación no puede ser negativo; Verifique", vbInformation: txtDiasLiquidacion.SetFocus: Exit Sub
  txtDiasLiquidacion.Text = FormatNumber(txtDiasLiquidacion.Text, 2)
End Sub
Private Sub txtDiasLiqVacacion_GotFocus()
  gdl_Procedure.MarcaGet txtDiasLiqVacacion
End Sub
Private Sub txtDiasLiqVacacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiasLiqVacacion_Validate(Cancel As Boolean)
  txtDiasLiqVacacion.Text = IIf(Not IsNumeric(txtDiasLiqVacacion.Text), 0, txtDiasLiqVacacion.Text)
  If CDec(txtDiasLiqVacacion.Text) < 0 Then MsgBox "Dias de Liquidación de Vacaciones no puede ser negativo; Verifique", vbInformation: txtDiasLiqVacacion.SetFocus: Exit Sub
  txtDiasLiqVacacion.Text = FormatNumber(txtDiasLiqVacacion.Text, 3)
End Sub
Private Sub txtDiaSuspension_GotFocus()
  gdl_Procedure.MarcaGet txtDiaSuspension
End Sub
Private Sub txtDiaSuspension_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaSuspension_Validate(Cancel As Boolean)
  txtDiaSuspension.Text = IIf(Not IsNumeric(txtDiaSuspension.Text), 0, txtDiaSuspension.Text)
  If CDec(txtDiaSuspension.Text) < 0 Then MsgBox "Dias suspensión no puede ser negativo; Verifique", vbInformation: txtDiaSuspension.SetFocus: Exit Sub
  txtDiaSuspension.Text = FormatNumber(txtDiaSuspension.Text, 2)
End Sub
Private Sub txtDiaTrabajo_GotFocus()
  gdl_Procedure.MarcaGet txtDiaTrabajo
End Sub
Private Sub txtDiaTrabajo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtDiaTrabajo_Validate(Cancel As Boolean)
  txtDiaTrabajo.Text = IIf(Not IsNumeric(txtDiaTrabajo.Text), 0, txtDiaTrabajo.Text)
  If CDec(txtDiaTrabajo.Text) < 0 Then MsgBox "Dias Trabajados no puede ser negativo; Verifique", vbInformation: txtDiaTrabajo.SetFocus: Exit Sub
  txtDiaTrabajo.Text = FormatNumber(txtDiaTrabajo.Text, 2)
End Sub
Private Sub txtEnfermedad_GotFocus()
  gdl_Procedure.MarcaGet txtEnfermedad
End Sub
Private Sub txtEnfermedad_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtEnfermedad_Validate(Cancel As Boolean)
  txtEnfermedad.Text = IIf(Not IsNumeric(txtEnfermedad.Text), 0, txtEnfermedad.Text)
  If CDec(txtEnfermedad.Text) < 0 Then MsgBox "Dias enfermedad no puede ser negativo; Verifique", vbInformation: txtEnfermedad.SetFocus: Exit Sub
  txtEnfermedad.Text = FormatNumber(txtEnfermedad.Text, 2)
End Sub
Private Sub txtFallece_GotFocus()
  gdl_Procedure.MarcaGet txtFallece
End Sub
Private Sub txtFallece_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtFallece_Validate(Cancel As Boolean)
  txtFallece.Text = IIf(Not IsNumeric(txtFallece.Text), 0, txtFallece.Text)
  If CDec(txtFallece.Text) < 0 Then MsgBox "Permiso por Fallecimiento no puede ser negativo; Verifique", vbInformation: txtFallece.SetFocus: Exit Sub
  txtFallece.Text = FormatNumber(txtFallece.Text, 2)
End Sub
Private Sub txtFaltas_GotFocus()
  gdl_Procedure.MarcaGet txtFaltas
End Sub
Private Sub txtFaltas_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtFaltas_Validate(Cancel As Boolean)
  txtFaltas.Text = IIf(Not IsNumeric(txtFaltas.Text), 0, txtFaltas.Text)
  If CDec(txtFaltas.Text) < 0 Then MsgBox "Faltas no puede ser negativo; Verifique", vbInformation: txtFaltas.SetFocus: Exit Sub
  txtFaltas.Text = FormatNumber(txtFaltas.Text, 2)
End Sub
Private Sub txtHoramedio_GotFocus()
  gdl_Procedure.MarcaGet txtHoraMedio
End Sub
Private Sub txtHoramedio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHoramedio_Validate(Cancel As Boolean)
  txtHoraMedio.Text = IIf(Not IsNumeric(txtHoraMedio.Text), 0, txtHoraMedio.Text)
  If CDec(txtHoraMedio.Text) < 0 Then MsgBox "Horas Medio Tiempo no puede ser negativo; Verifique", vbInformation: txtHoraMedio.SetFocus: Exit Sub
  txtHoraMedio.Text = FormatNumber(txtHoraMedio.Text, 2)
End Sub
Private Sub txtHoraNormal_GotFocus()
  gdl_Procedure.MarcaGet txtHoraNormal
End Sub
Private Sub txtHoraNormal_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHoraNormal_Validate(Cancel As Boolean)
  txtHoraNormal.Text = IIf(Not IsNumeric(txtHoraNormal.Text), 0, txtHoraNormal.Text)
  If CDec(txtHoraNormal.Text) < 0 Then MsgBox "Horas Normales no puede ser negativo; Verifique", vbInformation: txtHoraNormal.SetFocus: Exit Sub
  txtHoraNormal.Text = FormatNumber(txtHoraNormal.Text, 2)
End Sub
Private Sub txtHoraParcial_GotFocus()
  gdl_Procedure.MarcaGet txtHoraParcial
End Sub
Private Sub txtHoraParcial_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHoraParcial_Validate(Cancel As Boolean)
  txtHoraParcial.Text = IIf(Not IsNumeric(txtHoraParcial.Text), 0, txtHoraMedio.Text)
  If CDec(txtHoraParcial.Text) < 0 Then MsgBox "Horas Tiempo Parcial no puede ser negativo; Verifique", vbInformation: txtHoraParcial.SetFocus: Exit Sub
  txtHoraParcial.Text = FormatNumber(txtHoraParcial.Text, 2)
End Sub
Private Sub txtHrEspecial_GotFocus()
  gdl_Procedure.MarcaGet txtHrEspecial
End Sub
Private Sub txtHrEspecial_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHrEspecial_Validate(Cancel As Boolean)
  txtHrEspecial.Text = IIf(Not IsNumeric(txtHrEspecial.Text), 0, txtHrEspecial.Text)
  If CDec(txtHrEspecial.Text) < 0 Then MsgBox "Horas Especial no puede ser negativo; Verifique", vbInformation: txtHrEspecial.SetFocus: Exit Sub
  txtHrEspecial.Text = FormatNumber(txtHrEspecial.Text, 2)
End Sub
Private Sub txtHrExtraDoble_GotFocus()
  gdl_Procedure.MarcaGet txtHrExtraDoble
End Sub
Private Sub txtHrExtraDoble_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHrExtraDoble_Validate(Cancel As Boolean)
  txtHrExtraDoble.Text = IIf(Not IsNumeric(txtHrExtraDoble.Text), 0, txtHrExtraDoble.Text)
  If CDec(txtHrExtraDoble.Text) < 0 Then MsgBox "Horas Extras Dobles no puede ser negativo; Verifique", vbInformation: txtHrExtraDoble.SetFocus: Exit Sub
  txtHrExtraDoble.Text = FormatNumber(txtHrExtraDoble.Text, 2)
End Sub
Private Sub txtHrExtraSimple_GotFocus()
  gdl_Procedure.MarcaGet txtHrExtraSimple
End Sub
Private Sub txtHrExtraSimple_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHrExtraSimple_Validate(Cancel As Boolean)
  txtHrExtraSimple.Text = IIf(Not IsNumeric(txtHrExtraSimple.Text), 0, txtHrExtraSimple.Text)
  If CDec(txtHrExtraSimple.Text) < 0 Then MsgBox "Horas Extras Simples no puede ser negativo; Verifique", vbInformation: txtHrExtraSimple.SetFocus: Exit Sub
  txtHrExtraSimple.Text = FormatNumber(txtHrExtraSimple.Text, 2)
End Sub
Private Sub txtHrNocturno_GotFocus()
  gdl_Procedure.MarcaGet txtHrNocturno
End Sub
Private Sub txtHrNocturno_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtHrNocturno_Validate(Cancel As Boolean)
  txtHrNocturno.Text = IIf(Not IsNumeric(txtHrNocturno.Text), 0, txtHrNocturno.Text)
  If CDec(txtHrNocturno.Text) < 0 Then MsgBox "Horas Extras Nocturno no puede ser negativo; Verifique", vbInformation: txtHrNocturno.SetFocus: Exit Sub
  txtHrNocturno.Text = FormatNumber(txtHrNocturno.Text, 2)
End Sub
Private Sub txtLicencia_GotFocus()
  gdl_Procedure.MarcaGet txtLicencia
End Sub
Private Sub txtLicencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtLicencia_Validate(Cancel As Boolean)
  txtLicencia.Text = IIf(Not IsNumeric(txtLicencia.Text), 0, txtLicencia.Text)
  If CDec(txtLicencia.Text) < 0 Then MsgBox "Dias Licencia no puede ser negativo; Verifique", vbInformation: txtLicencia.SetFocus: Exit Sub
  txtLicencia.Text = FormatNumber(txtLicencia.Text, 2)
End Sub
Private Sub txtMotivo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtMotivo(Index)
End Sub
Private Sub txtMotivo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtMotivo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtMotivo_LostFocus(Index As Integer)
  lblHelp(Index).Caption = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtMotivo(Index).Text, "TS"), 30)
End Sub
Private Sub txtObservacion_GotFocus()
  gdl_Procedure.MarcaGet txtOpcional
End Sub
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtOpcional_GotFocus()
  gdl_Procedure.MarcaGet txtOpcional
End Sub
Private Sub txtOpcional_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtOpcional_Validate(Cancel As Boolean)
  txtOpcional.Text = IIf(Not IsNumeric(txtOpcional.Text), 0, txtOpcional.Text)
  If CDec(txtOpcional.Text) < 0 Then MsgBox "Dato opcional no puede ser negativo; Verifique", vbInformation: txtOpcional.SetFocus: Exit Sub
  txtOpcional.Text = FormatNumber(txtOpcional.Text, 2)
End Sub
Private Sub txtPaternidad_GotFocus()
  gdl_Procedure.MarcaGet txtpaternidad
End Sub
Private Sub txtPaternidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtPaternidad_Validate(Cancel As Boolean)
  txtpaternidad.Text = IIf(Not IsNumeric(txtpaternidad.Text), 0, txtpaternidad.Text)
  If CDec(txtpaternidad.Text) < 0 Then MsgBox "Dias Paternidad no puede ser negativo; Verifique", vbInformation: txtpaternidad.SetFocus: Exit Sub
  txtpaternidad.Text = FormatNumber(txtpaternidad.Text, 2)
End Sub
Private Sub txtPrePostNatal_GotFocus()
  gdl_Procedure.MarcaGet txtPrePostNatal
End Sub
Private Sub txtPrePostNatal_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtPrePostNatal_Validate(Cancel As Boolean)
  txtPrePostNatal.Text = IIf(Not IsNumeric(txtPrePostNatal.Text), 0, txtPrePostNatal.Text)
  If CDec(txtPrePostNatal.Text) < 0 Then MsgBox "Dias Pre-Post Natal no puede ser negativo; Verifique", vbInformation: txtPrePostNatal.SetFocus: Exit Sub
  txtPrePostNatal.Text = FormatNumber(txtPrePostNatal.Text, 2)
End Sub
Private Sub txtTardanzas_GotFocus()
  gdl_Procedure.MarcaGet txtTardanzas
End Sub
Private Sub txtTardanzas_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab
    KeyAscii = 0
  End If
End Sub
Private Sub txtTardanzas_Validate(Cancel As Boolean)
  txtTardanzas.Text = IIf(Not IsNumeric(txtTardanzas.Text), 0, txtTardanzas.Text)
  If CDec(txtTardanzas.Text) < 0 Then MsgBox "Horas Tarde no puede ser negativo; Verifique", vbInformation: txtTardanzas.SetFocus: Exit Sub
  txtTardanzas.Text = FormatNumber(txtTardanzas.Text, 2)
End Sub
Private Sub txtVacaciones_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtVacaciones(Index)
End Sub
Private Sub txtVacaciones_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtVacaciones_Validate(Index As Integer, Cancel As Boolean)
  txtVacaciones(Index).Text = IIf(Not IsNumeric(txtVacaciones(Index).Text), 0, txtVacaciones(Index).Text)
  If CDec(txtVacaciones(Index).Text) < 0 Then MsgBox "Dias Vacaciones no puede ser negativo; Verifique", vbInformation: txtVacaciones(Index).SetFocus: Exit Sub
  txtVacaciones(Index).Text = FormatNumber(txtVacaciones(Index).Text, 2)
End Sub
Private Sub txtVacaVencida_GotFocus()
  gdl_Procedure.MarcaGet txtVacaVencida
End Sub
Private Sub txtVacaVencida_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtVacaVencida_Validate(Cancel As Boolean)
  txtVacaVencida.Text = IIf(Not IsNumeric(txtVacaVencida.Text), 0, txtVacaVencida.Text)
  If CDec(txtVacaVencida.Text) < 0 Then MsgBox "Dias Vacaciones Vencidas no puede ser negativo; Verifique", vbInformation: txtVacaVencida.SetFocus: Exit Sub
  txtVacaVencida.Text = FormatNumber(txtVacaVencida.Text, 2)
End Sub
