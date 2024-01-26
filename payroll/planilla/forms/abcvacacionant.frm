VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcVacacionAnt 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7845
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   7845
      _Version        =   65536
      _ExtentX        =   13838
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
         Left            =   6975
         TabIndex        =   47
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
         Picture         =   "abcvacacionant.frx":0000
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6585
         TabIndex        =   48
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
         Picture         =   "abcvacacionant.frx":001C
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
         TabIndex        =   49
         Top             =   120
         Width           =   5205
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   50
      Top             =   5430
      Width           =   7845
      _Version        =   65536
      _ExtentX        =   13838
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
         Left            =   4695
         TabIndex        =   54
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
         Picture         =   "abcvacacionant.frx":0038
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   53
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
         Picture         =   "abcvacacionant.frx":0054
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   52
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
         Picture         =   "abcvacacionant.frx":0070
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   51
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
         Picture         =   "abcvacacionant.frx":008C
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4785
      Index           =   0
      Left            =   7050
      TabIndex        =   55
      Top             =   585
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8440
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
         TabIndex        =   56
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
         TabIndex        =   57
         Tag             =   "0"
         Top             =   600
         Visible         =   0   'False
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
         Picture         =   "abcvacacionant.frx":00A8
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   58
         Tag             =   "0"
         Top             =   1230
         Visible         =   0   'False
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
         Picture         =   "abcvacacionant.frx":00C4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   59
         Tag             =   "0"
         Top             =   1860
         Visible         =   0   'False
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
         Picture         =   "abcvacacionant.frx":00E0
      End
   End
   Begin Threed.SSFrame frmRegister 
      Height          =   4860
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   8572
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         ItemData        =   "abcvacacionant.frx":00FC
         Left            =   1440
         List            =   "abcvacacionant.frx":00FE
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1770
         Width           =   1485
      End
      Begin VB.ComboBox cmbAnyo 
         Height          =   315
         ItemData        =   "abcvacacionant.frx":0100
         Left            =   1440
         List            =   "abcvacacionant.frx":0102
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1485
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         ItemData        =   "abcvacacionant.frx":0104
         Left            =   1440
         List            =   "abcvacacionant.frx":0106
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1365
         Width           =   1485
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "abcvacacionant.frx":0108
         Left            =   5250
         List            =   "abcvacacionant.frx":010A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   1485
      End
      Begin VB.TextBox txtDias 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1020
         Width           =   765
      End
      Begin TabDlg.SSTab tabPago 
         Height          =   2340
         Left            =   105
         TabIndex        =   17
         Top             =   2355
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   4128
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   12582912
         TabCaption(0)   =   "Importes"
         TabPicture(0)   =   "abcvacacionant.frx":010C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "shpCuadro(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblDato(60)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblDato(18)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblDato(19)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "shpCuadro(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblDato(70)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblDato(8)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblDato(9)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtImporte(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtImporte(2)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtImporte(4)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtImporte(1)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtImporte(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtImporte(5)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "Cuentas"
         TabPicture(1)   =   "abcvacacionant.frx":0128
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "shpCuadro(2)"
         Tab(1).Control(1)=   "lblHelp(0)"
         Tab(1).Control(2)=   "lblDato(10)"
         Tab(1).Control(3)=   "lblHelp(1)"
         Tab(1).Control(4)=   "lblDato(11)"
         Tab(1).Control(5)=   "shpCuadro(3)"
         Tab(1).Control(6)=   "lblHelp(2)"
         Tab(1).Control(7)=   "lblDato(12)"
         Tab(1).Control(8)=   "lblHelp(3)"
         Tab(1).Control(9)=   "lblDato(13)"
         Tab(1).Control(10)=   "cmdHelp(3)"
         Tab(1).Control(11)=   "cmdHelp(2)"
         Tab(1).Control(12)=   "cmdHelp(1)"
         Tab(1).Control(13)=   "cmdHelp(0)"
         Tab(1).Control(14)=   "txtCuenta(0)"
         Tab(1).Control(15)=   "txtCuenta(1)"
         Tab(1).Control(16)=   "txtCuenta(2)"
         Tab(1).Control(17)=   "txtCuenta(3)"
         Tab(1).ControlCount=   18
         Begin VB.TextBox txtCuenta 
            Height          =   285
            Index           =   3
            Left            =   -73665
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   1755
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            Height          =   285
            Index           =   2
            Left            =   -73665
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   1425
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            Height          =   285
            Index           =   1
            Left            =   -73665
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   870
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            Height          =   285
            Index           =   0
            Left            =   -73665
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   540
            Width           =   1200
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Index           =   5
            Left            =   5130
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   1575
            Width           =   1245
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Index           =   3
            Left            =   5130
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1185
            Width           =   1245
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Index           =   1
            Left            =   5130
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   825
            Width           =   1245
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Index           =   4
            Left            =   1860
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1575
            Width           =   1245
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Index           =   2
            Left            =   1860
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   1185
            Width           =   1245
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Index           =   0
            Left            =   1860
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   825
            Width           =   1245
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   0
            Left            =   -72405
            TabIndex        =   38
            Top             =   540
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
            Left            =   -72405
            TabIndex        =   40
            Top             =   870
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
            Index           =   2
            Left            =   -72405
            TabIndex        =   42
            Top             =   1425
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
            Index           =   3
            Left            =   -72405
            TabIndex        =   44
            Top             =   1755
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Haber ME :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   -74760
            TabIndex        =   36
            Top             =   1785
            Width           =   1000
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
            Left            =   -72045
            TabIndex        =   45
            Top             =   1800
            Width           =   195
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Debe ME :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   -74760
            TabIndex        =   34
            Top             =   1455
            Width           =   1000
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
            Left            =   -72045
            TabIndex        =   43
            Top             =   1470
            Width           =   195
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   900
            Index           =   3
            Left            =   -74910
            Shape           =   4  'Rounded Rectangle
            Top             =   1290
            Width           =   6525
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Haber MN :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   -74760
            TabIndex        =   32
            Top             =   900
            Width           =   1000
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
            Left            =   -72045
            TabIndex        =   41
            Top             =   915
            Width           =   195
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Debe MN :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   -74760
            TabIndex        =   30
            Top             =   570
            Width           =   1000
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
            Left            =   -72045
            TabIndex        =   39
            Top             =   585
            Width           =   195
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   900
            Index           =   2
            Left            =   -74910
            Shape           =   4  'Rounded Rectangle
            Top             =   405
            Width           =   6525
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provisión ME :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   3510
            TabIndex        =   28
            Top             =   1605
            Width           =   1530
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Acumulado ME :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   3510
            TabIndex        =   26
            Top             =   1215
            Width           =   1530
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remuneración ME :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   70
            Left            =   3510
            TabIndex        =   24
            Top             =   855
            Width           =   1530
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   1305
            Index           =   1
            Left            =   3360
            Shape           =   4  'Rounded Rectangle
            Top             =   690
            Width           =   3285
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provisión MN :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   22
            Top             =   1605
            Width           =   1530
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Acumulado MN :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   20
            Top             =   1215
            Width           =   1530
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remuneración MN :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   60
            Left            =   240
            TabIndex        =   18
            Top             =   855
            Width           =   1530
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   1305
            Index           =   0
            Left            =   90
            Shape           =   4  'Rounded Rectangle
            Top             =   690
            Width           =   3285
         End
      End
      Begin MSComCtl2.DTPicker dtpFechas 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Top             =   690
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Format          =   113377281
         CurrentDate     =   37515
      End
      Begin MSComCtl2.DTPicker dtpFechas 
         Height          =   285
         Index           =   1
         Left            =   5250
         TabIndex        =   8
         Top             =   690
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Format          =   113377281
         CurrentDate     =   37515
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   285
         Left            =   5250
         TabIndex        =   14
         Top             =   1365
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
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
         BackStyle       =   0  'Transparent
         Caption         =   "F. Cancelacion :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   4020
         TabIndex        =   13
         Top             =   1395
         Width           =   1155
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   15
         Top             =   1785
         Width           =   1140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   4020
         TabIndex        =   7
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   11
         Top             =   1395
         Width           =   1140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes :"
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
         Height          =   195
         Index           =   1
         Left            =   4020
         TabIndex        =   3
         Top             =   375
         Width           =   1140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Año :"
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
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   5
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dias :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   9
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00800000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1995
         Index           =   4
         Left            =   105
         Shape           =   4  'Rounded Rectangle
         Top             =   210
         Width           =   6750
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2280
      TabIndex        =   60
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
Attribute VB_Name = "fAbcVacacionAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, n_Cuonter As Integer        ' Indice para bucle, y parametro de codigo
Private s_Registro As String, s_ParCodigo As String     ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Sub ShowScreen()
  
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditCombo "PK", cmbAnyo, -1, Me.Tag, False
    gdl_Procedure.EditCombo "PK", cmbMes, -1, Me.Tag, False
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txtDias, CInt(0), Me.Tag, False, 4, vbRightJustify
    gdl_Procedure.EditCombo "AT", cmbMoneda, -1, Me.Tag, False
    gdl_Procedure.EditMask "AT", mskFecha, "", Me.Tag, False, "##/##/####"
    gdl_Procedure.EditCombo "AT", cmbEstado, 0, Me.Tag, False
    
    gdl_Procedure.EditText "AT", txtImporte(0), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(1), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(2), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(3), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(4), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(5), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    
    gdl_Procedure.EditText "AT", txtCuenta(0), "", Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_debmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), "", Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_habmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(2), "", Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_debme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(3), "", Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_habme.DefinedSize
  Else
    n_Cuonter = CLng(o_PvsVacaConsul.dcaRegistro.Recordset!pdoano)
    n_Cuonter = IIf((Val(ps_Anyo) - n_Cuonter) >= 0, (20 - Abs(Val(ps_Anyo) - n_Cuonter)), (20 + Abs(Val(ps_Anyo) - n_Cuonter)))
    gdl_Procedure.EditCombo "PK", cmbAnyo, n_Cuonter, Me.Tag, True
    n_Cuonter = CInt(o_PvsVacaConsul.dcaRegistro.Recordset!pdomes)
    gdl_Procedure.EditCombo "PK", cmbMes, (n_Cuonter - 1), Me.Tag, True
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), o_PvsVacaConsul.dcaRegistro.Recordset!fechaini, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), o_PvsVacaConsul.dcaRegistro.Recordset!fechafin, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txtDias, CInt(o_PvsVacaConsul.dcaRegistro.Recordset!numerodias), Me.Tag, False, 4, vbRightJustify
    n_Cuonter = IIf(o_PvsVacaConsul.dcaRegistro.Recordset!codmon = "N", s_Estado_Ina, s_Estado_Act)
    gdl_Procedure.EditCombo "AT", cmbMoneda, n_Cuonter, Me.Tag, True
    gdl_Procedure.EditMask "AT", mskFecha, IIf(IsNull(o_PvsVacaConsul.dcaRegistro.Recordset!fechacan), "", o_PvsVacaConsul.dcaRegistro.Recordset!fechacan), Me.Tag, True, "##/##/####"
    n_Cuonter = CInt(o_PvsVacaConsul.dcaRegistro.Recordset!estadodet)
    gdl_Procedure.EditCombo "AT", cmbEstado, n_Cuonter, Me.Tag, False
    
    gdl_Procedure.EditText "AT", txtImporte(0), FormatNumber(o_PvsVacaConsul.dcaRegistro.Recordset!remunera_mn, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(1), FormatNumber(o_PvsVacaConsul.dcaRegistro.Recordset!remunera_me, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(2), FormatNumber(o_PvsVacaConsul.dcaRegistro.Recordset!imporpvsacu_mn, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(3), FormatNumber(o_PvsVacaConsul.dcaRegistro.Recordset!imporpvsacu_me, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(4), FormatNumber(o_PvsVacaConsul.dcaRegistro.Recordset!importepvs_mn, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtImporte(5), FormatNumber(o_PvsVacaConsul.dcaRegistro.Recordset!importepvs_me, 2), Me.Tag, False, 18, vbRightJustify
    
    gdl_Procedure.EditText "AT", txtCuenta(0), gdl_Funcion.aTexto(o_PvsVacaConsul.dcaRegistro.Recordset!codcta_debmn), Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_debmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), gdl_Funcion.aTexto(o_PvsVacaConsul.dcaRegistro.Recordset!codcta_habmn), Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_habmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(2), gdl_Funcion.aTexto(o_PvsVacaConsul.dcaRegistro.Recordset!codcta_debme), Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_debme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(3), gdl_Funcion.aTexto(o_PvsVacaConsul.dcaRegistro.Recordset!codcta_habme), Me.Tag, False, o_PvsVacaConsul.dcaRegistro.Recordset!codcta_habme.DefinedSize
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(0), "CU")
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(1), "CU")
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(2), "CU")
  lblHelp(3) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(3), "CU")

End Sub
Private Sub cmdAction_Click(Index As Integer)
  
  ' Valido que el peiodo no se eencuentre procesado
  If (cmbEstado.ListIndex = s_Estado_Blq And Index <> 0) Then Beep: MsgBox "Resgistro No se puede Actualizar se encuentra Cancelado", vbExclamation: Me.Tag = s_MdoData_Vis: Exit Sub
  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    cmbAnyo.SetFocus
  ElseIf Index = 2 Then
   dtpFechas(0).SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & cmbAnyo.Text & "-" & cmbMes.Text & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim(cmbAnyo.Text) & Trim(Left(cmbMes, 2))
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpvs", "codpsn", "pdopvs", "pdoano", "pdomes")
    a_Valores = Array(ps_ClsPlanilla, o_PvsVacaConsul.dcaRegistro.Recordset!codpvs, o_PvsVacaConsul.dcaRegistro.Recordset!codpsn, Trim(cmbAnyo.Text), Left(cmbMes.Text, 2))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plpvsvacaciondet", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl o_PvsVacaConsul.dcaRegistro, o_PvsVacaConsul.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((o_PvsVacaConsul.dcaRegistro.Recordset.EOF And o_PvsVacaConsul.dcaRegistro.Recordset.BOF) Or o_PvsVacaConsul.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      o_PvsVacaConsul.dcaRegistro.Recordset.Find ("cPrimaryKey >= '" & s_Registro & "'")
      If o_PvsVacaConsul.dcaRegistro.Recordset.EOF Then o_PvsVacaConsul.dcaRegistro.Recordset.MoveLast
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
   Case 0: o_PvsVacaConsul.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not o_PvsVacaConsul.dcaRegistro.Recordset.BOF Then o_PvsVacaConsul.dcaRegistro.Recordset.MovePrevious
           If o_PvsVacaConsul.dcaRegistro.Recordset.BOF Then o_PvsVacaConsul.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not o_PvsVacaConsul.dcaRegistro.Recordset.EOF Then o_PvsVacaConsul.dcaRegistro.Recordset.MoveNext
           If o_PvsVacaConsul.dcaRegistro.Recordset.EOF Then o_PvsVacaConsul.dcaRegistro.Recordset.MoveLast
   Case 3: o_PvsVacaConsul.dcaRegistro.Recordset.MoveLast
  End Select
  
End Sub
Private Sub cmdUpdate_Click()

'  ' Realizo las validaciones de los campos a actualizar
'  If Not (cmbAnyo.Text = Left(aPrimaryKey(2), 4) Or cmbAnyo.Text = Right(aPrimaryKey(2), 4)) Then Beep: MsgBox "Año  de " & lblTitle & " No valido", vbExclamation: cmbAnyo.SetFocus: Exit Sub
'  If cmbmes.Text = "" Then Beep: MsgBox "Debe Ingresar el Mes " & lblTitle, vbExclamation: cmbmes.SetFocus: Exit Sub
'  If Not (dtpFechas(0) >= aPrimaryKey(3) And dtpFechas(0) <= aPrimaryKey(4)) Then Beep: MsgBox "Fecha Inicial no valida", vbExclamation: dtpFechas(0).SetFocus: Exit Sub
'  If Not (dtpFechas(1) >= aPrimaryKey(3) And dtpFechas(1) <= aPrimaryKey(4)) Then Beep: MsgBox "Fecha Final no valida", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
'  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha de termino debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
'  If CInt(txtDias) = 0 Then Beep: MsgBox "Debe Ingresar los dias " & lblTitle, vbExclamation: txtDias.SetFocus: Exit Sub
'  If cmbMoneda = "" Then Beep: MsgBox "Debe Ingresar Moneda " & lblTitle, vbExclamation: cmbMoneda.SetFocus: Exit Sub
'  If txtImporte(0) = "" Then Beep: MsgBox "Debe Ingresar la Remuneracion en MN " & lblTitle, vbExclamation: txtImporte(0).SetFocus: Exit Sub
'  If txtImporte(1) = "" Then Beep: MsgBox "Debe Ingresar la Remuneracion en ME " & lblTitle, vbExclamation: txtImporte(1).SetFocus: Exit Sub
'  If txtImporte(2) = "" Then Beep: MsgBox "Debe Ingresar la Importe Acumulado en MN " & lblTitle, vbExclamation: txtImporte(2).SetFocus: Exit Sub
'  If txtImporte(3) = "" Then Beep: MsgBox "Debe Ingresar la Importe Acumulado en ME " & lblTitle, vbExclamation: txtImporte(3).SetFocus: Exit Sub
'  If txtImporte(4) = "" Then Beep: MsgBox "Debe Ingresar la Importe del Mes en MN " & lblTitle, vbExclamation: txtImporte(4).SetFocus: Exit Sub
'  If txtImporte(5) = "" Then Beep: MsgBox "Debe Ingresar la Importe del Mes en ME " & lblTitle, vbExclamation: txtImporte(5).SetFocus: Exit Sub
'
'  If txtCuenta(0) = "" Then Beep: MsgBox "Debe Ingresar Cuenta Debe MN " & lblTitle, vbExclamation: txtCuenta(0).SetFocus: Exit Sub
'  If lblHelp(0) = "???" Then Beep: MsgBox "Cuenta no valida; Verificar", vbExclamation: txtCuenta(0).SetFocus: Exit Sub
'  If txtCuenta(1) = "" Then Beep: MsgBox "Debe Ingresar Cuenta Haber MN " & lblTitle, vbExclamation: txtCuenta(1).SetFocus: Exit Sub
'  If lblHelp(1) = "???" Then Beep: MsgBox "Cuenta no valida; Verificar", vbExclamation: txtCuenta(1).SetFocus: Exit Sub
'  If txtCuenta(2) = "" Then Beep: MsgBox "Debe Ingresar Cuenta Debe ME " & lblTitle, vbExclamation: txtCuenta(2).SetFocus: Exit Sub
'  If lblHelp(2) = "???" Then Beep: MsgBox "Cuenta no valida; Verificar", vbExclamation: txtCuenta(2).SetFocus: Exit Sub
'  If txtCuenta(3) = "" Then Beep: MsgBox "Debe Ingresar Cuenta Haber ME " & lblTitle, vbExclamation: txtCuenta(3).SetFocus: Exit Sub
'  If lblHelp(3) = "???" Then Beep: MsgBox "Cuenta no valida; Verificar", vbExclamation: txtCuenta(3).SetFocus: Exit Sub
'  If cmbEstado = "" Then Beep: MsgBox "Debe Ingresar Estado " & lblTitle, vbExclamation: cmbEstado.SetFocus: Exit Sub
'
'  ' Coloco el puntero en espera
'  gdl_Procedure.PunteroEnEspera
'  ' Capturo el registro a actualizar
'  s_Registro = Trim(cmbAnyo.Text) & Left(Trim(cmbmes.Text), 2)
'
'  ' Creo los arreglos para la actualización
'  a_Campos = Array("codcls", "codpvs", "codpsn", "pdopvs", "pdoano", "pdomes", "fechaini", "fechafin", "numerodias", "codmon", "remunera_mn", "remunera_me", "imporpvsacu_mn", "imporpvsacu_me", "importepvs_mn", "importepvs_me", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme", "fechacan", "estadodet", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
'  a_Valores = Array(ps_ClsPlanilla, aPrimaryKey(0), aPrimaryKey(1), aPrimaryKey(2), Trim(cmbAnyo.Text), _
'                    Left(cmbmes.Text, 2), Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), CInt(txtDias.Text), _
'                    Choose(cmbMoneda.ListIndex + 1, s_Codmon_mn, s_Codmon_me), CDec(txtImporte(0).Text), CDec(txtImporte(1).Text), CDec(txtImporte(2).Text), CDec(txtImporte(3).Text), _
'                    CDec(txtImporte(4).Text), CDec(txtImporte(5).Text), Trim(txtCuenta(0).Text), Trim(txtCuenta(1).Text), Trim(txtCuenta(2).Text), _
'                    Trim(txtCuenta(3).Text), Format(mskFecha.Text, s_FmtFechMysql_0), cmbEstado.ListIndex, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
'  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
'                  TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, _
'                  TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA)
'  a_Where = Array("codcls", "codpvs", "codpsn", "pdopvs", "pdoano", "pdomes")
'
'  '[ Inicio la conexión a la base de datos ]
'  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
'
'  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
'  ' Realizo el proceso de actualización de los registros
'  If Me.Tag = s_MdoData_Ins Then
'    If Not Records_Ins("plpvsvacaciondet", a_Campos, a_Valores, a_Tipos) Then GoTo Error
'  Else
'    If Not Records_Upd("plpvsvacaciondet", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
'  End If
'  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
'
'  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
'  ' Refresco el ado control y la grilla
'  gdl_Procedure.RefreshAdoControl o_PvsVacaConsul.dcaRegistro, o_PvsVacaConsul.tdbRegistro, lblTitle
'  ' Ubico el registro ingresado o actualizado
'  o_PvsVacaConsul.dcaRegistro.Recordset.Find ("cPrimaryKey='" & s_Registro & "'")
'  ' si es actualización pasa al modo visualización
'  If Me.Tag = s_MdoData_Upd Then
'    cmdCancel_Click
'  Else
'    ShowScreen
'    cmbAnyo.SetFocus
'  End If
'  GoTo Finalizar
'
'Error:
'  gdl_Conexion.CancelaTransaccion
'Finalizar:
'  ' Coloco el puntero en normal
'  gdl_Procedure.PunteroNormal
'  '[ Finalizo la conexión a la base de datos ]
'  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click s_Estado_Act
End Sub
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Cuonter = 0 To 3: cmdMove(n_Cuonter).Visible = (Me.Tag = s_MdoData_Vis): Next n_Cuonter
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
   
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdHelp(1).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdHelp(2).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdHelp(3).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)

End Sub
Private Sub Form_Load()
  
  'Establece posición y titulo del formulario
  Me.Height = 6345: Me.Width = 7965
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización de Provisión de Vacaciones"
  lblTitle = "Provisión de Vacaciones"
  n_IndexHelp = -1
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operación del registro
  Me.Tag = o_PvsVacaConsul.Tag
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Cuonter = 0 To 2
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "anadir", "borrar", "modifica")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Cuonter
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Cuonter = 0 To 3
    aElemento(n_Cuonter, 1) = Choose(n_Cuonter + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Cuonter, 2) = Choose(n_Cuonter + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Cuonter
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (o_PvsVacaConsul.dcaRegistro.Recordset.EOF Or o_PvsVacaConsul.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = o_PvsVacaConsul.dcaRegistro.Recordset!pdoano
  
  ' Configuro los listados, datos adicionales
  For n_Cuonter = (Val(ps_Anyo) - 20) To (Val(ps_Anyo) + 20): cmbAnyo.AddItem Format(n_Cuonter, "0000"): Next n_Cuonter
  For n_Cuonter = 1 To 12: cmbMes.AddItem Choose(n_Cuonter, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Cuonter
  ' Adiciono los tipos de monedas
  For n_Cuonter = 0 To 1: cmbMoneda.AddItem Choose(n_Cuonter + 1, s_Codmon_mn_Nom, s_Codmon_me_Nom): Next n_Cuonter
  For n_Cuonter = 0 To 2: cmbEstado.AddItem Choose(n_Cuonter + 1, "Pendiente", "Provisionado", "Cancelado"): Next n_Cuonter
  
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Planilla", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub cmdHelp_Click(Index As Integer)

  Dim s_Conexion As String, s_CenCosto As String
  
  s_SqlHelp = ""
  s_Conexion = ps_StrgConnec & ps_DataBase
  Select Case Index
    Case 0, 1, 2, 3  ' Cuenta contable
    tdbHelp.Columns(0).DataField = "codcta": tdbHelp.Columns(1).DataField = "detcta"
    tdbHelp.Caption = "Cuenta Contable"
    s_Sql = gdl_Funcion.HelpTablas("cta", tdbHelp.Columns(0).DataField, ps_CodEmpresa, "")
    s_Conexion = ps_StrgConnec & ps_DaBasCon
  End Select
  Set porstHelp = OpenRecordset(s_Conexion, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (frmRegister.Top + (IIf(Index < 2, cmdHelp(Index).Top, 840) + (cmdHelp(Index).Height / 2)))
  tdbHelp.Left = IIf(Index < 2, 0, 1) + frmRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
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
   Case 0, 1, 2, 3  ' Cuenta contable
    txtCuenta(n_IndexHelp) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCuenta(n_IndexHelp).SetFocus
  End Select
 
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  Dim s_Conexion As String, s_CenCosto As String
  
  ' Recupero la información ordenada
  s_Conexion = ps_StrgConnec & ps_DataBase
  Select Case n_IndexHelp
    Case 0, 1, 2, 3 ' Cuenta contable
    s_Sql = gdl_Funcion.HelpTablas("cta", tdbHelp.Columns(ColIndex).DataField, ps_CodEmpresa, "")
    s_Conexion = ps_StrgConnec & ps_DaBasCon
  End Select
  Set porstHelp = OpenRecordset(s_Conexion, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtCuenta_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCuenta(Index)
End Sub
Private Sub txtCuenta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 3
End Sub
Private Sub txtCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCuenta_LostFocus(Index As Integer)
  lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(Index), "CU")
End Sub
Private Sub txtImporte_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtImporte(Index)
End Sub
Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    If CDec(txtImporte(Index)) <= 0 Then
      Beep
      MsgBox "Debe Ingresar el Valor de la " & lblTitle, vbExclamation
      txtImporte(Index).SetFocus
    Else
      SendKeys "{TAB}"
    End If
    KeyAscii = 0
  End If

End Sub
Private Sub txtimporte_Validate(Index As Integer, Cancel As Boolean)
  txtImporte(Index).Text = IIf(Not IsNumeric(txtImporte(Index).Text), 0, txtImporte(Index).Text)
  txtImporte(Index).Text = FormatNumber(CDec(txtImporte(Index).Text), 2)
End Sub

Private Sub mskFecha_GotFocus()
  gdl_Procedure.MarcaGet mskFecha
End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecha_Validate(Cancel As Boolean)
  If mskFecha.ClipText <> "" Then
    gdl_Funcion.ValidaFecha mskFecha, 1900
  End If
End Sub
