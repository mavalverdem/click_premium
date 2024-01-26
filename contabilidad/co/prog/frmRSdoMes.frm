VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRSdoMes 
   Caption         =   "[Entidad]"
   ClientHeight    =   7410
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab dgrCabecera 
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   12621
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Cuentas"
      TabPicture(0)   =   "frmRSdoMes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTexto(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dgrMain"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fgrMeses"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Diarios"
      TabPicture(1)   =   "frmRSdoMes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMes(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dgrDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Comprobante"
      TabPicture(2)   =   "frmRSdoMes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCabecera"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Detalle"
      TabPicture(3)   =   "frmRSdoMes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraCabecera 
         Height          =   6420
         Left            =   -74775
         TabIndex        =   53
         Top             =   495
         Width           =   9855
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   1995
            Left            =   270
            TabIndex        =   62
            Top             =   435
            Width           =   9195
            Begin VB.TextBox txtLlave 
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
               Left            =   900
               TabIndex        =   65
               Top             =   405
               Width           =   520
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
               Index           =   9
               Left            =   885
               TabIndex        =   64
               Top             =   1425
               Width           =   6400
            End
            Begin VB.TextBox txtLlave 
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
               Index           =   1
               Left            =   7320
               TabIndex        =   63
               Top             =   405
               Width           =   735
            End
            Begin MSComCtl2.DTPicker dtpFehCpb 
               Height          =   315
               Left            =   900
               TabIndex        =   66
               Top             =   1065
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Format          =   20643841
               CurrentDate     =   37953
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000002&
               BorderWidth     =   2
               X1              =   180
               X2              =   8820
               Y1              =   900
               Y2              =   900
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Diario:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   2
               Left            =   225
               TabIndex        =   71
               Top             =   465
               Width           =   450
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Fecha :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   4
               Left            =   225
               TabIndex        =   70
               Top             =   1125
               Width           =   540
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Glosa:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   5
               Left            =   225
               TabIndex        =   69
               Top             =   1485
               Width           =   465
            End
            Begin VB.Label lblLlaveDeta 
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
               Left            =   1425
               TabIndex        =   68
               Top             =   405
               Width           =   3675
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "NºComprobante:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   3
               Left            =   5880
               TabIndex        =   67
               Top             =   465
               Width           =   1185
            End
         End
         Begin VB.Frame Frame3 
            ForeColor       =   &H80000002&
            Height          =   1095
            Left            =   135
            TabIndex        =   54
            Top             =   5175
            Width           =   9420
            Begin VB.TextBox txtDeta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
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
               Index           =   3
               Left            =   7440
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   645
               Width           =   1755
            End
            Begin VB.TextBox txtDeta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
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
               Left            =   5640
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   225
               Width           =   1755
            End
            Begin VB.TextBox txtDeta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
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
               Index           =   1
               Left            =   7440
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   225
               Width           =   1755
            End
            Begin VB.TextBox txtDeta 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
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
               Index           =   2
               Left            =   5640
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   645
               Width           =   1755
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Totales M.E. :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   7
               Left            =   4440
               TabIndex        =   60
               Top             =   690
               Width           =   960
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Totales M.N. :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   6
               Left            =   4440
               TabIndex        =   59
               Top             =   285
               Width           =   975
            End
         End
         Begin MSDataGridLib.DataGrid dgrDetalleCab 
            Height          =   2595
            Left            =   270
            TabIndex        =   61
            Top             =   2550
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   4577
            _Version        =   393216
            AllowUpdate     =   0   'False
            ForeColor       =   -2147483630
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label txtMes 
            AutoSize        =   -1  'True
            Caption         =   "TXTMES"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   1
            Left            =   320
            TabIndex        =   73
            Top             =   230
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   6420
         Left            =   -74775
         TabIndex        =   8
         Top             =   495
         Width           =   9855
         Begin VB.ComboBox CboTpoTCb 
            Height          =   315
            Left            =   2805
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   5265
            Width           =   915
         End
         Begin VB.TextBox txtDato 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Index           =   8
            Left            =   3885
            TabIndex        =   35
            Top             =   5265
            Width           =   735
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
            Index           =   7
            Left            =   1605
            TabIndex        =   34
            Top             =   4305
            Width           =   7020
         End
         Begin VB.ComboBox cboTpoMon 
            Height          =   315
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   5265
            Width           =   675
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "##.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   5265
            TabIndex        =   32
            Top             =   4905
            Width           =   1575
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Index           =   1
            Left            =   7065
            TabIndex        =   31
            Top             =   4905
            Width           =   1575
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Index           =   2
            Left            =   5265
            TabIndex        =   30
            Top             =   5265
            Width           =   1575
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Index           =   3
            Left            =   7065
            TabIndex        =   29
            Top             =   5265
            Width           =   1575
         End
         Begin VB.Frame fraDocumento 
            Caption         =   "Documento"
            ForeColor       =   &H80000002&
            Height          =   1980
            Left            =   885
            TabIndex        =   12
            Top             =   2160
            Width           =   7755
            Begin VB.OptionButton optTpoPvs 
               Caption         =   "&Otros"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   255
               Index           =   2
               Left            =   6255
               TabIndex        =   72
               Top             =   1530
               Width           =   1215
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
               Index           =   6
               Left            =   1320
               TabIndex        =   18
               Top             =   1380
               Width           =   1695
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
               Index           =   5
               Left            =   1740
               TabIndex        =   17
               Top             =   900
               Width           =   1155
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
               Index           =   4
               Left            =   1320
               TabIndex        =   16
               Top             =   900
               Width           =   435
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
               Index           =   3
               Left            =   1320
               TabIndex        =   15
               Top             =   360
               Width           =   315
            End
            Begin VB.OptionButton optTpoPvs 
               Caption         =   "&Cancelación"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   255
               Index           =   1
               Left            =   6240
               TabIndex        =   14
               Top             =   1155
               Width           =   1215
            End
            Begin VB.OptionButton optTpoPvs 
               Caption         =   "&Provisión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   255
               Index           =   0
               Left            =   6240
               TabIndex        =   13
               Top             =   795
               Width           =   975
            End
            Begin MSComCtl2.DTPicker dtpFeEDoc 
               Height          =   315
               Left            =   4320
               TabIndex        =   19
               Top             =   780
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               OLEDropMode     =   1
               Format          =   20643841
               CurrentDate     =   37959.8076041667
            End
            Begin MSComCtl2.DTPicker dtpFeVDoc 
               Height          =   315
               Left            =   4320
               TabIndex        =   20
               Top             =   1140
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   20643841
               CurrentDate     =   37977.3988773148
            End
            Begin MSComCtl2.DTPicker dtpFeRDoc 
               Height          =   315
               Left            =   4320
               TabIndex        =   21
               Top             =   1500
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   20643841
               CurrentDate     =   37977.3989467593
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "NºDocumento:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   13
               Left            =   285
               TabIndex        =   28
               Top             =   900
               Width           =   1035
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Referencia:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   15
               Left            =   480
               TabIndex        =   27
               Top             =   1380
               Width           =   840
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   12
               Left            =   120
               TabIndex        =   26
               Top             =   420
               Width           =   1200
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
               Index           =   3
               Left            =   1620
               TabIndex        =   25
               Top             =   360
               Width           =   4635
            End
            Begin VB.Label lblTexto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Fecha Emisión:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   14
               Left            =   3210
               TabIndex        =   24
               Top             =   825
               Width           =   1080
            End
            Begin VB.Label lblTexto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vencimiento:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   16
               Left            =   3360
               TabIndex        =   23
               Top             =   1170
               Width           =   930
            End
            Begin VB.Label lblTexto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Recepción:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   210
               Index           =   17
               Left            =   3480
               TabIndex        =   22
               Top             =   1530
               Width           =   810
            End
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
            Left            =   1740
            TabIndex        =   11
            Top             =   945
            Width           =   975
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
            Index           =   1
            Left            =   1740
            TabIndex        =   10
            Top             =   1305
            Width           =   615
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
            Index           =   2
            Left            =   1740
            TabIndex        =   9
            Top             =   1665
            Width           =   1275
         End
         Begin MSComCtl2.DTPicker dtpFehOpe 
            Height          =   315
            Left            =   2700
            TabIndex        =   37
            Top             =   585
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   37924.6695138889
         End
         Begin VB.Label txtMes 
            AutoSize        =   -1  'True
            Caption         =   "TXTMES"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   960
            TabIndex        =   74
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lblTexto 
            Caption         =   "Tipo de Cambio:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   20
            Left            =   2925
            TabIndex        =   51
            Top             =   4905
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   18
            Left            =   1005
            TabIndex        =   50
            Top             =   4305
            Width           =   465
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Mon. Func.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   19
            Left            =   1725
            TabIndex        =   49
            Top             =   4905
            Width           =   840
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Debe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   21
            Left            =   5865
            TabIndex        =   48
            Top             =   4665
            Width           =   375
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Haber"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   22
            Left            =   7605
            TabIndex        =   47
            Top             =   4665
            Width           =   435
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   " M.N."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   23
            Left            =   4845
            TabIndex        =   46
            Top             =   4965
            Width           =   360
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   " M.E."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   24
            Left            =   4845
            TabIndex        =   45
            Top             =   5325
            Width           =   345
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   9
            Left            =   990
            TabIndex        =   44
            Top             =   1005
            Width           =   555
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
            Left            =   2700
            TabIndex        =   43
            Top             =   945
            Width           =   5940
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
            Index           =   1
            Left            =   2340
            TabIndex        =   42
            Top             =   1305
            Width           =   3735
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "C.Costo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   10
            Left            =   990
            TabIndex        =   41
            Top             =   1365
            Width           =   615
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
            Index           =   2
            Left            =   3000
            TabIndex        =   40
            Top             =   1665
            Width           =   5625
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Auxiliar:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   11
            Left            =   990
            TabIndex        =   39
            Top             =   1725
            Width           =   585
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Operación:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   210
            Index           =   8
            Left            =   990
            TabIndex        =   38
            Top             =   645
            Width           =   1515
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgrMeses 
         Bindings        =   "frmRSdoMes.frx":0070
         Height          =   3990
         Left            =   180
         TabIndex        =   1
         Top             =   3015
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7038
         _Version        =   393216
         Rows            =   16
         Cols            =   9
         BackColorSel    =   -2147483640
      End
      Begin MSDataGridLib.DataGrid dgrMain 
         Height          =   2145
         Left            =   225
         TabIndex        =   4
         Top             =   495
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3784
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   6195
         Left            =   -74760
         TabIndex        =   7
         Top             =   720
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   10927
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label txtMes 
         AutoSize        =   -1  'True
         Caption         =   "TXTMES"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   -74460
         TabIndex        =   52
         Top             =   495
         Width           =   660
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "MONEDA EXTRANJERA"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   5790
         TabIndex        =   6
         Top             =   2835
         Width           =   1815
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "MONEDA NACIONAL"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   2430
         TabIndex        =   5
         Top             =   2835
         Width           =   1560
      End
      Begin VB.Label Label28 
         Caption         =   "Alimentos Por Asignar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   450
         Left            =   -74760
         TabIndex        =   3
         Top             =   4920
         Width           =   870
      End
      Begin VB.Label Label15 
         Caption         =   "Alimentos Asignados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   450
         Left            =   -74700
         TabIndex        =   2
         Top             =   2580
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmRSdoMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain_0 As ADODB.Recordset, _
       uorstMain_1 As ADODB.Recordset, _
       uorstMain_2 As ADODB.Recordset, _
       uorstMain_3 As ADODB.Recordset, _
       uorstMain_4 As ADODB.Recordset
Public uorstCODro As ADODB.Recordset, _
       uorstTGTCb As ADODB.Recordset, _
       uorstCOCta As ADODB.Recordset, _
       uorstCoCCo As ADODB.Recordset, _
       uorstTGAux As ADODB.Recordset, _
       uorstTGTDc As ADODB.Recordset, _
       uorstCOCpbDet As ADODB.Recordset, _
       uorstCOTCbMes As ADODB.Recordset
Private porstCOCta As ADODB.Recordset
Public usConnStrgSele_0 As String, _
       usConnStrgOrde_0 As String, _
       usConnStrgSele_1 As String, _
       usConnStrgWher_1 As String, _
       usConnStrgOrde_1 As String, _
       usConnStrgSele_2 As String, _
       usConnStrgWher_2 As String, _
       usConnStrgOrde_2 As String, _
       usConnStrgSele_3 As String, _
       usConnStrgWher_3 As String, _
       usConnStrgOrde_3 As String, _
       usConnStrgSele_4 As String, _
       usConnStrgWher_4 As String, _
       usConnStrgOrde_4 As String
Private pnColumnaOrd As Integer

'Dim WithEvents MRViewer As MRViewerObject
'
'Public udFecha As Date
'Public unCopias As Integer
'Public unMargenIzquierdo As Integer
'Public usDEstino As String
'Public usOrientacionRpt As String
'Public usOrientacionOri As String
'Private paOpciones As Variant
'Private pocnnMain As ADODB.Connection
'Private porstMRp As ADODB.Recordset

Private Sub Form_Load()
 
 '[Recordsets.
   usConnStrgSele_0 = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
   usConnStrgSele_0 = usConnStrgSele_0 & "FROM CoCta "
   usConnStrgSele_0 = usConnStrgSele_0 & "WHERE codemp='" & gsCodEmp & "' "
   usConnStrgSele_0 = usConnStrgSele_0 & "AND pdoano='" & gsAnoAct & "' "
   usConnStrgSele_0 = usConnStrgSele_0 & "AND TpoCta='" & TPOCTA_TRA & "' "
   usConnStrgOrde_0 = "ORDER BY CodCta"
   
   usConnStrgSele_1 = "SELECT CodCta, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD00_MN, AcuH00_MN, AcuD00_ME, AcuH00_ME, AcuD00_MN-AcuH00_MN AS cAcu00_MN, AcuD00_ME-AcuH00_ME AS cAcu00_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD01_MN, AcuH01_MN, AcuD01_ME, AcuH01_ME, AcuD01_MN-AcuH01_MN AS cAcu01_MN, AcuD01_ME-AcuH01_ME AS cAcu01_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD02_MN, AcuH02_MN, AcuD02_ME, AcuH02_ME, AcuD02_MN-AcuH02_MN AS cAcu02_MN, AcuD02_ME-AcuH02_ME AS cAcu02_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD03_MN, AcuH03_MN, AcuD03_ME, AcuH03_ME, AcuD03_MN-AcuH03_MN AS cAcu03_MN, AcuD03_ME-AcuH03_ME AS cAcu03_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD04_MN, AcuH04_MN, AcuD04_ME, AcuH04_ME, AcuD04_MN-AcuH04_MN AS cAcu04_MN, AcuD04_ME-AcuH04_ME AS cAcu04_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD05_MN, AcuH05_MN, AcuD05_ME, AcuH05_ME, AcuD05_MN-AcuH05_MN AS cAcu05_MN, AcuD05_ME-AcuH05_ME AS cAcu05_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD06_MN, AcuH06_MN, AcuD06_ME, AcuH06_ME, AcuD06_MN-AcuH06_MN AS cAcu06_MN, AcuD06_ME-AcuH06_ME AS cAcu06_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD07_MN, AcuH07_MN, AcuD07_ME, AcuH07_ME, AcuD07_MN-AcuH07_MN AS cAcu07_MN, AcuD07_ME-AcuH07_ME AS cAcu07_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD08_MN, AcuH08_MN, AcuD08_ME, AcuH08_ME, AcuD08_MN-AcuH08_MN AS cAcu08_MN, AcuD08_ME-AcuH08_ME AS cAcu08_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD09_MN, AcuH09_MN, AcuD09_ME, AcuH09_ME, AcuD09_MN-AcuH09_MN AS cAcu09_MN, AcuD09_ME-AcuH09_ME AS cAcu09_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD10_MN, AcuH10_MN, AcuD10_ME, AcuH10_ME, AcuD10_MN-AcuH10_MN AS cAcu10_MN, AcuD10_ME-AcuH10_ME AS cAcu10_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD11_MN, AcuH11_MN, AcuD11_ME, AcuH11_ME, AcuD11_MN-AcuH11_MN AS cAcu11_MN, AcuD11_ME-AcuH11_ME AS cAcu11_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD12_MN, AcuH12_MN, AcuD12_ME, AcuH12_ME, AcuD12_MN-AcuH12_MN AS cAcu12_MN, AcuD12_ME-AcuH12_ME AS cAcu12_ME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "AcuD13_MN, AcuH13_MN, AcuD13_ME, AcuH13_ME, AcuD13_MN-AcuH13_MN AS cAcu13_MN, AcuD13_ME-AcuH13_ME AS cAcu13_ME "
   usConnStrgSele_1 = usConnStrgSele_1 & "FROM COCtaAcu "
   usConnStrgWher_1 = "WHERE codemp='" & gsCodEmp & "' "
   usConnStrgWher_1 = usConnStrgWher_1 & "AND pdoano='" & gsAnoAct & "' "
   usConnStrgOrde_1 = "ORDER BY codcta"
   
   usConnStrgSele_2 = "SELECT a.CodCta, a.CodCCo, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, a.GloIte, "
   usConnStrgSele_2 = usConnStrgSele_2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END) AS cImpMNDebe, "
   usConnStrgSele_2 = usConnStrgSele_2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END) AS cImpMNHaber, "
   usConnStrgSele_2 = usConnStrgSele_2 & "(CASE a.TpoGnr WHEN " & TPOGNR_DRO & " THEN '" & TPOGNR_DRO_TXT & "' WHEN " & TPOGNR_CPR & " THEN '" & TPOGNR_CPR_TXT & "' WHEN " & TPOGNR_VTA & " THEN '" & TPOGNR_VTA_TXT & "' WHEN " & TPOGNR_HPR & " THEN '" & TPOGNR_HPR_TXT & "' WHEN " & TPOGNR_DST & " THEN '" & TPOGNR_DST_TXT & "' WHEN " & TPOGNR_DCA & " THEN '" & TPOGNR_DCA_TXT & "' WHEN " & TPOGNR_APE & " THEN '" & TPOGNR_APE_TXT & "' ELSE '" & TPOGNR_CIE_TXT & "' END) AS ccTpoGnr, "
   usConnStrgSele_2 = usConnStrgSele_2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END) AS cImpMEDebe, "
   usConnStrgSele_2 = usConnStrgSele_2 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END) AS cImpMEHaber, "
   usConnStrgSele_2 = usConnStrgSele_2 & "a.MesPvs, a.CodDro, a.NroCpb, a.NroIte, "
   usConnStrgSele_2 = usConnStrgSele_2 & "d.AbvTDc, a.TpoCtb, b.FehCpb, b.GloCpb, "
   usConnStrgSele_2 = usConnStrgSele_2 & Choose(gsIdioma, "c.DetDro", "c.DetDrox") & " AS DetDro "
   usConnStrgSele_2 = usConnStrgSele_2 & "FROM (((COCpbDet a "
   usConnStrgSele_2 = usConnStrgSele_2 & "LEFT JOIN COCpbCab b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.COdDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.MesPvs=b.MesPvs) "
   usConnStrgSele_2 = usConnStrgSele_2 & "LEFT JOIN CODro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.COdDro=c.CodDro) "
   usConnStrgSele_2 = usConnStrgSele_2 & "LEFT JOIN TgTdc d ON a.codemp=d.codemp AND a.COdTDc=d.CodTDc) "
   usConnStrgWher_2 = "WHERE a.codemp='" & gsCodEmp & "' "
   usConnStrgWher_2 = usConnStrgWher_2 & "AND a.pdoano='" & gsAnoAct & "' "
   usConnStrgOrde_2 = "ORDER BY a.NroIte"
   
   usConnStrgSele_3 = "SELECT a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.FehOpe, a.CodCCo, a.CodAux, a.CodTdc, b.AbvTDc, "
   usConnStrgSele_3 = usConnStrgSele_3 & "a.SerDoc, a.NroDoc, a.GloIte, a.TpoCtb, "
   usConnStrgSele_3 = usConnStrgSele_3 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END) AS cImpMNDebe, "
   usConnStrgSele_3 = usConnStrgSele_3 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END) AS cImpMNHaber, "
   usConnStrgSele_3 = usConnStrgSele_3 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END) AS cImpMEDebe, "
   usConnStrgSele_3 = usConnStrgSele_3 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END) AS cImpMEHaber "
   usConnStrgSele_3 = usConnStrgSele_3 & "FROM (COCpbDet a "
   usConnStrgSele_3 = usConnStrgSele_3 & "LEFT JOIN TgTDc b ON a.codemp=b.codemp AND a.COdTdc=b.CodTDc) "
   usConnStrgWher_3 = "WHERE a.codemp='" & gsCodEmp & "' "
   usConnStrgWher_3 = usConnStrgWher_3 & "AND a.pdoano='" & gsAnoAct & "' "
   usConnStrgOrde_3 = "ORDER BY a.CodDro, a.NroCpb, a.NroIte"
   
   usConnStrgSele_4 = "SELECT a.CodDro, a.NroCpb, a.MesPvs, a.CodCta, a.FehOpe, a.CodCCo, a.CodAux, a.CodTdc, "
   usConnStrgSele_4 = usConnStrgSele_4 & "a.SerDoc, a.NroDoc, a.RefDoc, a.GloIte, a.ImpTCb, a.FeEDoc, a.FeVDoc, a.FeRDoc, a.TpoCtb, "
   usConnStrgSele_4 = usConnStrgSele_4 & "a.TpoMon, a.TpoTcb, a.TpoPvs, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, " & Choose(gsIdioma, "c.DetCCo", "c.DetCCox") & " AS DetCCo, d.RazAux, " & Choose(gsIdioma, "e.DetTDc", "e.DetTDcx") & " AS DetTDc, "
   usConnStrgSele_4 = usConnStrgSele_4 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END) AS cImpMNDebe, "
   usConnStrgSele_4 = usConnStrgSele_4 & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END) AS cImpMEDebe, "
   usConnStrgSele_4 = usConnStrgSele_4 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END) AS cImpMNHaber, "
   usConnStrgSele_4 = usConnStrgSele_4 & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END) AS cImpMEHaber "
   usConnStrgSele_4 = usConnStrgSele_4 & "FROM ((((COCpbDet a "
   usConnStrgSele_4 = usConnStrgSele_4 & "LEFT JOIN COCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
   usConnStrgSele_4 = usConnStrgSele_4 & "LEFT JOIN COCCo c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCCo=c.CodCCo) "
   usConnStrgSele_4 = usConnStrgSele_4 & "LEFT JOIN TgAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux) "
   usConnStrgSele_4 = usConnStrgSele_4 & "LEFT JOIN TgTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
   usConnStrgWher_4 = "WHERE a.codemp='" & gsCodEmp & "' "
   usConnStrgWher_4 = usConnStrgWher_4 & "AND a.pdoano='" & gsAnoAct & "' "
   usConnStrgOrde_4 = "ORDER BY a.CodDro, a.NroCpb, a.MesPvs"
   
   Set uocnnMain = New ADODB.Connection
   Set uorstMain_0 = New ADODB.Recordset
   Set uorstMain_1 = New ADODB.Recordset
   Set uorstMain_2 = New ADODB.Recordset
   Set uorstMain_3 = New ADODB.Recordset
   Set uorstMain_4 = New ADODB.Recordset
   Set uorstCODro = New ADODB.Recordset
   Set uorstTGTCb = New ADODB.Recordset
   Set uorstCOCta = New ADODB.Recordset
   Set uorstCoCCo = New ADODB.Recordset
   Set uorstTGAux = New ADODB.Recordset
   Set uorstTGTDc = New ADODB.Recordset
   Set uorstCOCpbDet = New ADODB.Recordset
   Set uorstCOTCbMes = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With uorstMain_0
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_0 & usConnStrgOrde_0
'     .CursorLocation = adUseClient   'Es el Default.
        
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open (usConnStrgSele_0 & usConnStrgOrde_0)
      .Properties("Unique Table").Value = "COCpbCab"
   End With
   With uorstMain_1
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_1 & usConnStrgWher_1 & usConnStrgOrde_1
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstMain_2
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_2 & usConnStrgWher_2 & usConnStrgOrde_2
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstMain_3
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_3 & usConnStrgWher_3 & usConnStrgOrde_3
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
   End With

   With uorstMain_4
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_4 & usConnStrgWher_4 & usConnStrgOrde_4
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   
   With uorstCODro
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodDro, DetDro, Cpb" & gsMesAct & " "
      .Source = .Source & "FROM CODro "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro) > 2"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstTGTCb
      .ActiveConnection = uocnnMain
      .Source = "SELECT FehTCb, ImpTCb_Cpr, ImpTCb_Vta "
      .Source = .Source & "FROM TGTCb "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND Month(FehTCb)=" & Val(gsMesAct) & " "
      .Source = .Source & "AND Year(FehTCb)=" & Val(gsAnoAct)
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCOCta
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta, TpoTCb, "
      .Source = .Source & "IndAjd, IndCCo, IndDoc, CodCta_AjD_Deb, CodCta_AjD_Hab "
      .Source = .Source & "FROM COCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND TpoCta=" & TPOCTA_TRA & " "
      .Source = .Source & "AND EstCta='" & ESTCTA_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCoCCo
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS DetCCo "
      .Source = .Source & "FROM COCCo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND EstCCo='" & ESTCCO_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstTGAux
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TGAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND EstAux='" & ESTAUX_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstTGTDc
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodTDc, DetTDc "
      .Source = .Source & "FROM TGTDc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCOCpbDet
      .ActiveConnection = uocnnMain
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
   End With
   With uorstCOTCbMes
      .ActiveConnection = uocnnMain
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
   End With

    With cboTpoMon
       .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
       .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
    End With
    With cboTpoTCb
       .AddItem TPOTCB_VTA_TXT, TPOTCB_VTA_IND
       .AddItem TPOTCB_CPR_TXT, TPOTCB_CPR_IND
    End With
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(25, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "MONEDA NACIONAL", "MONEDA EXTRANJERA", "Diario :", "Nº Comprobante :", "Fecha :", "Glosa :", "Totales M.N. :", "Totales M.E. :", "Fecha de Operación :", "Cuenta :", "C.Costo :", "Auxiliar :", "Tipo Documento :", "Nº Documento :", "Fecha Emisión :", "Referencia :", "Vencimiento :", "Recepción", "Glosa :", "Mon. Func. :", "Tipo de Cambio :", "Debe", "Haber", "M.N.", "M.E.")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "NATIONAL CURRENCY", "FOREIGN CURRENCY", "Journal :", "Nº Voucher :", "Date :", "Gloss :", "Totals N.C. :", "Totals F.C. :", "Operation Date :", "Account :", "C.Center :", "Auxiliary :", "Type Document :", "Nº Document :", "Issue Date :", "Reference :", "Due Date :", "Reception :", "Gloss :", "Func. Cur. :", "Rate of Exchange :", "Debit", "Credit", "N.C.", "F.C.")
  Next nElemento
  fraDocumento.Caption = Choose(gsIdioma, "Documento", "Document")
  optTpoPvs(0).Caption = Choose(gsIdioma, "&Provisión", "&Provision")
  optTpoPvs(1).Caption = Choose(gsIdioma, "&Cancelación", "&Cancelation")
  optTpoPvs(2).Caption = Choose(gsIdioma, "&Otros", "&Others")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, False, aLabel
 ']
  dgrCabecera.TabCaption(0) = Choose(gsIdioma, "Cuentas", "Accounts")
  dgrCabecera.TabCaption(1) = Choose(gsIdioma, "Diarios", "Journals")
  dgrCabecera.TabCaption(2) = Choose(gsIdioma, "Comprobantes", "Vouchers")
  dgrCabecera.TabCaption(3) = Choose(gsIdioma, "Detalle", "Detail")

   dgrCabecera.Tab = 0
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain_0
End Sub

Private Sub Form_Activate()
   'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   ppDatosGrid
   ppDatosGridMeses
   ppDatosDetalle True
'   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   If uorstMain_2.State = adStateOpen Then uorstMain_2.Close
   If uorstMain_3.State = adStateOpen Then uorstMain_3.Close
   If uorstMain_4.State = adStateOpen Then uorstMain_4.Close
   uorstMain_0.Close
   uocnnMain.Close
   Set uorstMain_2 = Nothing
   Set uorstMain_3 = Nothing
   Set uorstMain_4 = Nothing
   Set uorstCOTCbMes = Nothing
   Set uorstMain_0 = Nothing
   Set uocnnMain = Nothing
End Sub

Public Sub ppLimpiaVariables()
   txtLlave(0).Text = "": txtLlave(1).Text = ""
   lblLlaveDeta(0).Caption = ""
   lblDatoDeta(0).Caption = "": lblDatoDeta(1).Caption = "": lblDatoDeta(2).Caption = "": lblDatoDeta(3).Caption = ""
   txtDato(0).Text = "": txtDato(1).Text = "": txtDato(2).Text = "": txtDato(3).Text = "": txtDato(4).Text = "": txtDato(5).Text = "": txtDato(6).Text = "": txtDato(7).Text = "": txtDato(8).Text = "": txtDato(9).Text = ""
   txtDeta(0).Text = "": txtDeta(1).Text = "": txtDeta(2).Text = "": txtDeta(3).Text = ""
   txtImporte(0).Text = "0.00": txtImporte(1).Text = "0.00": txtImporte(2).Text = "0.00": txtImporte(3).Text = "0.00"
   dtpFehCpb.Value = Format(Date, "dd/mm/yyyy")
   dtpFehOpe.Value = Format(Date, "dd/mm/yyyy")
   dtpFeEDoc.Value = Format(Date, "dd/mm/yyyy")
   dtpFeVDoc.Value = Format(Date, "dd/mm/yyyy")
   dtpFeRDoc.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub dgrDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If LastCol <> 0 Then ppDatosCabecera
End Sub

Private Sub dgrDetalleCab_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If LastCol <> 0 Then ppDetalleDiario
End Sub

Private Sub dgrMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   ppDatosDetalle True
End Sub

Private Sub fgrMeses_RowColChange()
   ppDatosDetalle False
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyHome
      uorstMain_0.MoveFirst
   Case vbKeyEnd
      uorstMain_0.MoveLast
   End Select
End Sub

Public Sub ppDatosGridMeses()
   Dim dnNum    As Integer
        
   With fgrMeses
      .Clear
      For dnNum = 0 To .cols - 1
         Select Case dnNum
         Case 0: .ColWidth(dnNum) = 1500
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "     Meses\Saldos", "     Months\Balances")
         Case 1: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Debe", "Debit")
         Case 2: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Haber", "Credit")
         Case 3: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Saldo", "Balance")
         Case 4: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Acumulado", "Accrued")
         Case 5: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Debe", "Debit")
         Case 6: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Haber", "Credit")
         Case 7: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Saldo", "Balance")
         Case 8: .ColWidth(dnNum) = 1000
                 .ColAlignment(dnNum) = 7
                 .TextMatrix(0, dnNum) = Choose(gsIdioma, "Acumulado", "Accrued")
         End Select
      Next dnNum
      
      For dnNum = 0 To .Rows - 1
         Select Case dnNum
         Case 1:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Apertura", "Opening")
         Case 2:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Enero", "January")
         Case 3:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Febrero", "February")
         Case 4:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Marzo", "March")
         Case 5:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Abril", "April")
         Case 6:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Mayo", "May")
         Case 7:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Junio", "June")
         Case 8:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Julio", "July")
         Case 9:  .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Agosto", "August")
         Case 10: .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Setiembre", "September")
         Case 11: .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Octubre", "October")
         Case 12: .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Noviembre", "November")
         Case 13: .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Diciembre", "December")
         Case 14: .TextMatrix(dnNum, 0) = Choose(gsIdioma, "Cierre", "Closing")
         Case 15: .TextMatrix(dnNum, 0) = Choose(gsIdioma, "TOTALES", "TOTALS")
         End Select
      Next dnNum
      .row = Val(gsMesAct) + 2
      .SetFocus
   End With
End Sub

Public Sub ppDatosDetalle(ByVal Ind As Boolean)
   Dim dncol As Integer
   Dim dnfil As Integer
   Dim dnDebeMN As Double, dnDebeME As Double, dnHaberMN As Double, dnHaberME As Double
   Dim cTotalMN As Double, cTotalME As Double, cAcumMN As Double, cAcumME As Double
   
   dnDebeMN = 0#: dnDebeME = 0#: dnHaberMN = 0#: dnHaberME = 0#
   cTotalMN = 0#: cTotalME = 0#: cAcumMN = 0#: cAcumME = 0#
   If Ind Then
      With uorstMain_1
         If .RecordCount > 0 Then
            .MoveFirst
            .Find "CodCta='" & uorstMain_0!codcta & "'"
            If Not .EOF Then
               For dnfil = 0 To fgrMeses.Rows - 3
                  With fgrMeses
                     .TextMatrix(dnfil + 1, 1) = Format(uorstMain_1.Fields("AcuD" & Format(dnfil, "00") & "_MN"), FORMATO_NUM_1)
                     dnDebeMN = dnDebeMN + Format(.TextMatrix(dnfil + 1, 1), FORMATO_NUM_2)
                     .TextMatrix(dnfil + 1, 2) = Format(uorstMain_1.Fields("AcuH" & Format(dnfil, "00") & "_MN"), FORMATO_NUM_1)
                     dnHaberMN = dnHaberMN + Format(.TextMatrix(dnfil + 1, 2), FORMATO_NUM_2)
                     .TextMatrix(dnfil + 1, 3) = Format(uorstMain_1.Fields("cAcu" & Format(dnfil, "00") & "_MN"), FORMATO_NUM_1)
                     cTotalMN = cTotalMN + Format(.TextMatrix(dnfil + 1, 3), FORMATO_NUM_2)
                     cAcumMN = cAcumMN + Format(.TextMatrix(dnfil + 1, 3), FORMATO_NUM_2)
                     .TextMatrix(dnfil + 1, 4) = Format(cAcumMN, FORMATO_NUM_1)
                     .TextMatrix(dnfil + 1, 5) = Format(uorstMain_1.Fields("AcuD" & Format(dnfil, "00") & "_ME"), FORMATO_NUM_1)
                     dnDebeME = dnDebeME + Format(.TextMatrix(dnfil + 1, 5), FORMATO_NUM_2)
                     .TextMatrix(dnfil + 1, 6) = Format(uorstMain_1.Fields("AcuH" & Format(dnfil, "00") & "_ME"), FORMATO_NUM_1)
                     dnHaberME = dnHaberME + Format(.TextMatrix(dnfil + 1, 6), FORMATO_NUM_2)
                     .TextMatrix(dnfil + 1, 7) = Format(uorstMain_1.Fields("cAcu" & Format(dnfil, "00") & "_ME"), FORMATO_NUM_1)
                     cTotalME = cTotalME + Format(.TextMatrix(dnfil + 1, 7), FORMATO_NUM_2)
                     cAcumME = cAcumME + Format(.TextMatrix(dnfil + 1, 7), FORMATO_NUM_2)
                     .TextMatrix(dnfil + 1, 8) = Format(cAcumME, FORMATO_NUM_1)
                  End With
               Next dnfil
               With fgrMeses
                  .TextMatrix(15, 1) = Format(dnDebeMN, FORMATO_NUM_1)
                  .TextMatrix(15, 2) = Format(dnHaberMN, FORMATO_NUM_1)
                  .TextMatrix(15, 3) = Format(cTotalMN, FORMATO_NUM_1)
                  .TextMatrix(15, 5) = Format(dnDebeME, FORMATO_NUM_1)
                  .TextMatrix(15, 6) = Format(dnHaberMN, FORMATO_NUM_1)
                  .TextMatrix(15, 7) = Format(cTotalME, FORMATO_NUM_1)
               End With
            Else
               For dnfil = 0 To fgrMeses.Rows - 3
                  With fgrMeses
                     .TextMatrix(dnfil + 1, 1) = "0.00"
                     .TextMatrix(dnfil + 1, 2) = "0.00"
                     .TextMatrix(dnfil + 1, 3) = "0.00"
                     .TextMatrix(dnfil + 1, 4) = "0.00"
                     .TextMatrix(dnfil + 1, 5) = "0.00"
                     .TextMatrix(dnfil + 1, 6) = "0.00"
                     .TextMatrix(dnfil + 1, 7) = "0.00"
                     .TextMatrix(dnfil + 1, 8) = "0.00"
                  End With
               Next dnfil
               With fgrMeses
                  .TextMatrix(15, 1) = "0.00"
                  .TextMatrix(15, 2) = "0.00"
                  .TextMatrix(15, 3) = "0.00"
                  .TextMatrix(15, 5) = "0.00"
                  .TextMatrix(15, 6) = "0.00"
                  .TextMatrix(15, 7) = "0.00"
               End With
            End If
         Else
            For dnfil = 0 To fgrMeses.Rows - 3
               With fgrMeses
                  .TextMatrix(dnfil + 1, 1) = "0.00"
                  .TextMatrix(dnfil + 1, 2) = "0.00"
                  .TextMatrix(dnfil + 1, 3) = "0.00"
                  .TextMatrix(dnfil + 1, 4) = "0.00"
                  .TextMatrix(dnfil + 1, 5) = "0.00"
                  .TextMatrix(dnfil + 1, 6) = "0.00"
                  .TextMatrix(dnfil + 1, 7) = "0.00"
                  .TextMatrix(dnfil + 1, 8) = "0.00"
               End With
            Next dnfil
            With fgrMeses
               .TextMatrix(15, 1) = "0.00"
               .TextMatrix(15, 2) = "0.00"
               .TextMatrix(15, 3) = "0.00"
               .TextMatrix(15, 5) = "0.00"
               .TextMatrix(15, 6) = "0.00"
               .TextMatrix(15, 7) = "0.00"
            End With
         End If
     End With
   End If
    
'[para grid Detalle
    
'    usConnStrgWher_3 = ""
    usConnStrgWher_3 = "WHERE a.codemp='" & gsCodEmp & "' "
    usConnStrgWher_3 = usConnStrgWher_3 & "AND a.pdoano='" & gsAnoAct & "' "
    usConnStrgWher_3 = usConnStrgWher_3 & "AND a.MesPvs='" & Format(fgrMeses.row - 1, "00") & "' "
    usConnStrgWher_3 = usConnStrgWher_3 & "AND a.CodCta='" & dgrMain.Columns(0) & "' "
    With uorstMain_3
       If .State = adStateOpen Then .Close
       .Source = usConnStrgSele_3 & usConnStrgWher_3 & usConnStrgOrde_3
       .Open
    End With
    
    dgrDetalle.MarqueeStyle = dbgHighlightRow
    Set dgrDetalle.DataSource = uorstMain_3
    
    ppDatosGridDetalle
    
    If uorstMain_3.RecordCount > 0 Then
       ppDatosCabecera
    Else
       If uorstMain_2.State = adStateOpen Then uorstMain_2.Close
       If uorstMain_4.State = adStateOpen Then uorstMain_4.Close
       ppLimpiaVariables
    End If
    
    txtMes(0).Caption = Choose(gsIdioma, "Mes analizado: ", "Analyzed Month: ") & fgrMeses.TextMatrix(fgrMeses.row, 0) & Choose(gsIdioma, "   Cuenta : ", "   Account : ") & uorstMain_0!codcta
    txtMes(1).Caption = Choose(gsIdioma, "Mes analizado: ", "Analyzed Month: ") & fgrMeses.TextMatrix(fgrMeses.row, 0)
    txtMes(2).Caption = Choose(gsIdioma, "Mes analizado: ", "Analyzed Month: ") & fgrMeses.TextMatrix(fgrMeses.row, 0)
']
End Sub

Public Sub ppDatosCabecera()
   On Error GoTo ERR13

'[para grid Cabecera y Detalle
    
   usConnStrgWher_2 = ""
   usConnStrgWher_2 = "WHERE a.codemp='" & gsCodEmp & "' "
   usConnStrgWher_2 = usConnStrgWher_2 & "AND a.MesPvs='" & Format(dgrDetalle.Columns(0), "00") & "' "
   usConnStrgWher_2 = usConnStrgWher_2 & "AND a.pdoano='" & gsAnoAct & "' "
   usConnStrgWher_2 = usConnStrgWher_2 & "AND a.CodDro='" & dgrDetalle.Columns(2) & "' "
   usConnStrgWher_2 = usConnStrgWher_2 & "AND a.NroCpb='" & dgrDetalle.Columns(3) & "' "
                 
   '------------Llenando La pestaña de comprobante------------
   With uorstMain_2
      If .State = adStateOpen Then .Close
      .Source = usConnStrgSele_2 & usConnStrgWher_2 & usConnStrgOrde_2
      .Open
      txtLlave(0).Text = .Fields!coddro
      txtLlave(1).Text = .Fields!NroCpb
      txtDato(9).Text = .Fields!glocpb
      lblLlaveDeta(0).Caption = .Fields!DetDro
      txtDeta(0).Text = Format(.Fields!cImpMNDebe, FORMATO_NUM_1)
      txtDeta(1).Text = Format(.Fields!cImpMNHaber, FORMATO_NUM_1)
      txtDeta(2).Text = Format(.Fields!cImpMEDebe, FORMATO_NUM_1)
      txtDeta(3).Text = Format(.Fields!cImpMEHaber, FORMATO_NUM_1)
   End With
    
   dgrDetalleCab.MarqueeStyle = dbgHighlightRow
   Set dgrDetalleCab.DataSource = uorstMain_2
   
   txtDeta(0).Text = 0#
   txtDeta(2).Text = 0#
   txtDeta(1).Text = 0#
   txtDeta(3).Text = 0#

   If uorstMain_2.RecordCount > 0 Then
      uorstMain_2.MoveFirst
      Do
         txtDeta(0).Text = Format(CDec(txtDeta(0).Text) + uorstMain_2.Fields!cImpMNDebe, FORMATO_NUM_1)
         txtDeta(2).Text = Format(CDec(txtDeta(2).Text) + uorstMain_2.Fields!cImpMEDebe, FORMATO_NUM_1)
         txtDeta(1).Text = Format(CDec(txtDeta(1).Text) + uorstMain_2.Fields!cImpMNHaber, FORMATO_NUM_1)
         txtDeta(3).Text = Format(CDec(txtDeta(3).Text) + uorstMain_2.Fields!cImpMEHaber, FORMATO_NUM_1)
         uorstMain_2.MoveNext
      Loop Until uorstMain_2.EOF
      uorstMain_2.MoveFirst
   End If
      
   ppDatosDiario
']

ERR13:  Resume Next
End Sub

Public Sub ppDetalleDiario()
   On Error GoTo ERR13

  '-------------------Llenando Pestaña de Detalle-------------
    
'   usConnStrgWher_4 = ""
   usConnStrgWher_4 = "WHERE a.codemp='" & gsCodEmp & "' "
   usConnStrgWher_4 = usConnStrgWher_4 & "AND a.pdoano='" & gsAnoAct & "' "
   usConnStrgWher_4 = usConnStrgWher_4 & "AND a.MesPvs='" & Format(dgrDetalleCab.Columns(12), "00") & "' "
   usConnStrgWher_4 = usConnStrgWher_4 & "AND a.CodDro='" & dgrDetalleCab.Columns(13) & "' "
   usConnStrgWher_4 = usConnStrgWher_4 & "AND a.NroCpb='" & dgrDetalleCab.Columns(14) & "' "
   usConnStrgWher_4 = usConnStrgWher_4 & "AND a.NroIte = '" & dgrDetalleCab.Columns(15) & "' "
   With uorstMain_4
      .Close
      .Source = usConnStrgSele_4 & usConnStrgWher_4 & usConnStrgOrde_4
      .Open
      dtpFehOpe.Value = Format(.Fields!FehOpe, "dd/mm/yyyy")
      dtpFeEDoc.Value = Format(.Fields!FeEDoc, "dd/mm/yyyy")
      dtpFeVDoc.Value = Format(.Fields!fevdoc, "dd/mm/yyyy")
      dtpFeRDoc.Value = Format(.Fields!ferdoc, "dd/mm/yyyy")
      
      txtDato(0).Text = Trim("" & .Fields!codcta)
      txtDato(1).Text = Trim("" & .Fields!codcco)
      txtDato(2).Text = Trim("" & .Fields!CodAux)
      txtDato(3).Text = Trim("" & .Fields!CodTDc)
      txtDato(4).Text = Trim("" & .Fields!SerDoc)
      txtDato(5).Text = Trim("" & .Fields!NroDoc)
      txtDato(6).Text = Trim("" & .Fields!refdoc)
      txtDato(7).Text = Trim("" & .Fields!GloIte)
      txtDato(8).Text = Format(.Fields!ImpTCb, FORMATO_NUM_1)
      
      lblDatoDeta(0).Caption = Trim("" & .Fields!detcta)
      lblDatoDeta(1).Caption = Trim("" & .Fields!DetCCo)
      lblDatoDeta(2).Caption = Trim("" & .Fields!RazAux)
      lblDatoDeta(3).Caption = Trim("" & .Fields!dettdc)
      
      cboTpoMon.ListIndex = IIf(.Fields!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      cboTpoTCb.ListIndex = IIf(.Fields!TpoTcb = TPOTCB_VTA, TPOTCB_VTA_IND, TPOTCB_CPR_IND)
      
      optTpoPvs(0).Value = IIf(.Fields!TpoPvs = TPOPVS_PVS, True, False)
      optTpoPvs(1).Value = IIf(.Fields!TpoPvs = TPOPVS_CAN, True, False)
      optTpoPvs(2).Value = IIf(.Fields!TpoPvs = TPOPVS_OTR, True, False)
      
      txtImporte(0).Text = Format(.Fields!cImpMNDebe, FORMATO_NUM_1)
      txtImporte(1).Text = Format(.Fields!cImpMNHaber, FORMATO_NUM_1)
      txtImporte(2).Text = Format(.Fields!cImpMEDebe, FORMATO_NUM_1)
      txtImporte(3).Text = Format(.Fields!cImpMEHaber, FORMATO_NUM_1)
   End With
   
   Exit Sub
    
ERR13: Resume Next
End Sub

Public Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("CodCta").DefinedSize + 2)
        Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Descripción", "Description")
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("DetCta").DefinedSize + 2)
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "         Tipo de Moneda", "       Type of Currency")
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("cTpoMon").DefinedSize + 4)
            .Item(dnNum).Alignment = dbgCenter
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

Public Sub ppDatosGridDetalle()
   Dim dnNum As Integer
        
   With dgrDetalle.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 2
             .Item(dnNum).Caption = Choose(gsIdioma, "Diario", "Journal")
             .Item(dnNum).Width = 100 * (uorstMain_3.Fields("CodDro").DefinedSize + 2)
         Case 3
             .Item(dnNum).Caption = Choose(gsIdioma, "N.Comprob.", "NºVoucher")
             .Item(dnNum).Width = 100 * (uorstMain_3.Fields("NroCpb").DefinedSize + 4)
             .Item(dnNum).Alignment = dbgLeft
         Case 4
             .Item(dnNum).Caption = Choose(gsIdioma, "F.Opera.", "Opera. Date")
             .Item(dnNum).Width = 100 * (uorstMain_3.Fields("FehOpe").DefinedSize + 4)
             .Item(dnNum).Alignment = dbgCenter
         Case 8
             .Item(dnNum).Caption = "T.Doc"
             .Item(dnNum).Width = 100 * (uorstMain_3.Fields("AbvTDc").DefinedSize + 3)
             .Item(dnNum).Alignment = dbgCenter
         Case 9
             .Item(dnNum).Caption = "Ser.Doc."
             .Item(dnNum).Width = 100 * (uorstMain_3.Fields("SerDoc").DefinedSize + 4)
             .Item(dnNum).Alignment = dbgCenter
         Case 10
             .Item(dnNum).Caption = "Num.Doc."
             .Item(dnNum).Width = 100 * (uorstMain_3.Fields("NroDoc").DefinedSize + 2)
             .Item(dnNum).Alignment = dbgCenter
         Case 13
             .Item(dnNum).Caption = Choose(gsIdioma, "Debe", "Debit")
             .Item(dnNum).Width = 1000
             .Item(dnNum).Alignment = dbgRight
             .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 14
             .Item(dnNum).Caption = Choose(gsIdioma, "Haber", "Credit")
             .Item(dnNum).Width = 1000
             .Item(dnNum).Alignment = dbgRight
             .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 15
             .Item(dnNum).Caption = Choose(gsIdioma, "Debe", "Debit")
             .Item(dnNum).Width = 1000
             .Item(dnNum).Alignment = dbgRight
             .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 16
             .Item(dnNum).Caption = Choose(gsIdioma, "Haber", "Credit")
             .Item(dnNum).Width = 1000
             .Item(dnNum).Alignment = dbgRight
             .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case Else
             .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

Public Sub ppDatosDiario()
   Dim dnNum As Integer
        
   With dgrDetalleCab.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
            .Item(dnNum).Width = 80 * (uorstMain_2.Fields("CodCta").DefinedSize + 2)
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "C.Cto.", "C.Center")
            .Item(dnNum).Width = 100 * (uorstMain_2.Fields("CodCCo").DefinedSize + 2)
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 90 * (uorstMain_2.Fields("CodAux").DefinedSize + 2)
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "T.Doc.", "T.Doc.")
            .Item(dnNum).Width = 100 * (uorstMain_2.Fields("AbvTDc").DefinedSize + 2)
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "Serie", "Series")
            .Item(dnNum).Width = 100 * (uorstMain_2.Fields("SerDoc").DefinedSize + 2)
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Número", "Number")
            .Item(dnNum).Width = 1000
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
            .Item(dnNum).Width = 1100
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "Debe", "Debit")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 8
            .Item(dnNum).Caption = Choose(gsIdioma, "Haber", "Credit")
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 9
            .Item(dnNum).Caption = Choose(gsIdioma, "Tipo", "Type")
            .Item(dnNum).Width = 800
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub


Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
'   cmdNuevo.Enabled = taOpciones(0)
'   cmdEliminar.Enabled = taOpciones(1)
'   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

