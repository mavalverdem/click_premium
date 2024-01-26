VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fPrmCertifikCts 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5160
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "prmcerticts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6420
   Begin TabDlg.SSTab tabRegister 
      Height          =   3960
      Left            =   75
      TabIndex        =   30
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   6985
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
      TabCaption(0)   =   "Analisis"
      TabPicture(0)   =   "prmcerticts.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHelp(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblHelp(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblHelp(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblHelp(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDato(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblHelp(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDato(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDato(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblHelp(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdHelp(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdHelp(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdHelp(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdHelp(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdHelp(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdHelp(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRemunera(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtGratificacion"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtImporteCts"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtRemunera(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtImportePvs"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtRemunera(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Certificado"
      TabPicture(1)   =   "prmcerticts.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCalculoCts(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtCalculoCts(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtCalculoCts(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdHelp(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdHelp(8)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdHelp(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblDato(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblHelp(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblHelp(8)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblDato(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblHelp(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblDato(7)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
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
         Height          =   280
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   3225
         Width           =   980
      End
      Begin VB.TextBox txtImportePvs 
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
         Height          =   280
         Left            =   240
         TabIndex        =   13
         Top             =   2655
         Width           =   980
      End
      Begin VB.TextBox txtCalculoCts 
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
         Height          =   280
         Index           =   0
         Left            =   -74760
         TabIndex        =   19
         Top             =   405
         Width           =   980
      End
      Begin VB.TextBox txtCalculoCts 
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
         Height          =   280
         Index           =   2
         Left            =   -74760
         TabIndex        =   25
         Top             =   1515
         Width           =   980
      End
      Begin VB.TextBox txtCalculoCts 
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
         Height          =   280
         Index           =   1
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   980
      End
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
         Height          =   280
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   980
      End
      Begin VB.TextBox txtImporteCts 
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
         Height          =   280
         Left            =   240
         TabIndex        =   10
         Top             =   2085
         Width           =   980
      End
      Begin VB.TextBox txtGratificacion 
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
         Height          =   280
         Left            =   240
         TabIndex        =   7
         Top             =   1515
         Width           =   980
      End
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
         Height          =   280
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   405
         Width           =   980
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   1290
         TabIndex        =   31
         Top             =   405
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   2
         Left            =   1290
         TabIndex        =   33
         Top             =   1515
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   3
         Left            =   1290
         TabIndex        =   34
         Top             =   2085
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   1
         Left            =   1290
         TabIndex        =   32
         Top             =   960
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   6
         Left            =   -73710
         TabIndex        =   40
         Top             =   405
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   8
         Left            =   -73710
         TabIndex        =   42
         Top             =   1515
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   7
         Left            =   -73710
         TabIndex        =   41
         Top             =   960
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   4
         Left            =   1290
         TabIndex        =   37
         Top             =   2655
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   36
         Top             =   3225
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
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
         Left            =   1635
         TabIndex        =   17
         Top             =   3270
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Remuneració Percibida :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   15
         Top             =   2985
         Width           =   1845
      End
      Begin VB.Label lblDato 
         Caption         =   "Provisión del Mes :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   2415
         Width           =   1845
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
         Left            =   1650
         TabIndex        =   14
         Top             =   2700
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Importe por Dias :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   -74760
         TabIndex        =   18
         Top             =   165
         Width           =   1845
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
         Left            =   -73350
         TabIndex        =   20
         Top             =   450
         Width           =   195
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
         Left            =   -73350
         TabIndex        =   26
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Importe por Años :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   24
         Top             =   1275
         Width           =   1845
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
         Left            =   -73350
         TabIndex        =   23
         Top             =   1005
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Importe por Meses :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   -74760
         TabIndex        =   21
         Top             =   720
         Width           =   1845
      End
      Begin VB.Label lblDato 
         Caption         =   "Remuneració Promedio :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1850
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
         Left            =   1650
         TabIndex        =   5
         Top             =   1005
         Width           =   195
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
         Left            =   1650
         TabIndex        =   11
         Top             =   2130
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Importe C.T.S. :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1845
         Width           =   1850
      End
      Begin VB.Label lblDato 
         Caption         =   "Gratificación :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1275
         Width           =   1850
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
         Left            =   1650
         TabIndex        =   8
         Top             =   1560
         Width           =   195
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
         Left            =   1650
         TabIndex        =   2
         Top             =   450
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Remuneración Basica :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   165
         Width           =   1850
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
         Left            =   5790
         TabIndex        =   38
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
         Picture         =   "prmcerticts.frx":0044
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   5400
         TabIndex        =   39
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmcerticts.frx":0060
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
         Left            =   285
         TabIndex        =   28
         Top             =   120
         Width           =   4800
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   29
      Top             =   4650
      Width           =   6420
      _Version        =   65536
      _ExtentX        =   11324
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   35
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
Attribute VB_Name = "fPrmCertifikCts"
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
'[
Sub ShowScreen()
    
  ' Información de configuración
  s_Sql = "SELECT remubasicacts, remupromects, remugraticts, remutotalcts, remunepvscts, remupercibects, remudiascts, remumesescts, remuanoscts "
  s_Sql = s_Sql & "FROM plparametroafp "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    Me.Tag = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtRemunera(0), gdl_Funcion.aTexto(porstRecordset!remubasicacts), Me.Tag, False, porstRecordset!remubasicacts.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(1), gdl_Funcion.aTexto(porstRecordset!remupromects), Me.Tag, False, porstRecordset!remupromects.DefinedSize
    gdl_Procedure.EditText "AT", txtGratificacion, gdl_Funcion.aTexto(porstRecordset!remugraticts), Me.Tag, False, porstRecordset!remugraticts.DefinedSize
    gdl_Procedure.EditText "AT", txtImporteCts, gdl_Funcion.aTexto(porstRecordset!remutotalcts), Me.Tag, False, porstRecordset!remutotalcts.DefinedSize
    gdl_Procedure.EditText "AT", txtImportePvs, gdl_Funcion.aTexto(porstRecordset!remunepvscts), Me.Tag, False, porstRecordset!remutotalcts.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(2), gdl_Funcion.aTexto(porstRecordset!remupercibects), Me.Tag, False, porstRecordset!remupercibects.DefinedSize
    gdl_Procedure.EditText "AT", txtCalculoCts(0), gdl_Funcion.aTexto(porstRecordset!remudiascts), Me.Tag, False, porstRecordset!remudiascts.DefinedSize
    gdl_Procedure.EditText "AT", txtCalculoCts(1), gdl_Funcion.aTexto(porstRecordset!remumesescts), Me.Tag, False, porstRecordset!remumesescts.DefinedSize
    gdl_Procedure.EditText "AT", txtCalculoCts(2), gdl_Funcion.aTexto(porstRecordset!remuanoscts), Me.Tag, False, porstRecordset!remuanoscts.DefinedSize
  Else
    Me.Tag = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtRemunera(0), "", Me.Tag, False, porstRecordset!remubasicacts.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(1), "", Me.Tag, False, porstRecordset!remupromects.DefinedSize
    gdl_Procedure.EditText "AT", txtGratificacion, "", Me.Tag, False, porstRecordset!remugraticts.DefinedSize
    gdl_Procedure.EditText "AT", txtImporteCts, "", Me.Tag, False, porstRecordset!remutotalcts.DefinedSize
    gdl_Procedure.EditText "AT", txtImportePvs, "", Me.Tag, False, porstRecordset!remunepvscts.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(2), "", Me.Tag, False, porstRecordset!remupercibects.DefinedSize
    gdl_Procedure.EditText "AT", txtCalculoCts(0), "", Me.Tag, False, porstRecordset!remudiascts.DefinedSize
    gdl_Procedure.EditText "AT", txtCalculoCts(1), "", Me.Tag, False, porstRecordset!remumesescts.DefinedSize
    gdl_Procedure.EditText "AT", txtCalculoCts(2), "", Me.Tag, False, porstRecordset!remuanoscts.DefinedSize
  End If
  lblHelp(0).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(0).Text, "CP")
  lblHelp(1).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(1).Text, "CP")
  lblHelp(2).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtGratificacion.Text, "CP")
  lblHelp(3).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtImporteCts.Text, "CP")
  lblHelp(4).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtImportePvs.Text, "CP")
  lblHelp(5).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(2).Text, "CP")
  lblHelp(6).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtCalculoCts(0).Text, "CP")
  lblHelp(7).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtCalculoCts(1).Text, "CP")
  lblHelp(8).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtCalculoCts(2).Text, "CP")

End Sub
']
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0, 1, 5   ' Conceptos de planilla remuneracion
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Ingresos"
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
   Case 2, 3, 4    ' Conceptos de planilla gratificacion
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Ingresos"
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
   Case 6, 7, 8   ' Conceptos de planilla calculo cts
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Ingresos"
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + (Choose(Index + 1, cmdHelp(Index).Top, 750, 800, 850, 850, 850, cmdHelp(Index).Top, 750, 800) + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  
  ' Realizo las validaciones de los campos a actualizar
  If txtRemunera(0).Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Basica", vbExclamation: txtRemunera(0).SetFocus: Exit Sub
  If lblHelp(0).Caption = "???" Then Beep: MsgBox "Concepto de Remuneración Basica no es valido; Verificar", vbExclamation: txtRemunera(0).SetFocus: Exit Sub
  If txtRemunera(1).Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Promedio", vbExclamation: txtRemunera(1).SetFocus: Exit Sub
  If lblHelp(1).Caption = "???" Then Beep: MsgBox "Concepto de Remuneración Promedio no es valido; Verificar", vbExclamation: txtRemunera(1).SetFocus: Exit Sub
  If txtGratificacion.Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Gratificación", vbExclamation: txtGratificacion.SetFocus: Exit Sub
  If lblHelp(2).Caption = "???" Then Beep: MsgBox "Concepto de Gratificación no es valido; Verificar", vbExclamation: txtGratificacion.SetFocus: Exit Sub
  If txtImporteCts.Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Importe C.T.S.", vbExclamation: txtImporteCts.SetFocus: Exit Sub
  If lblHelp(3).Caption = "???" Then Beep: MsgBox "Concepto de Importe C.T.S. no es valido; Verificar", vbExclamation: txtImporteCts.SetFocus: Exit Sub
  If txtImportePvs.Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Provisión C.T.S.", vbExclamation: txtImportePvs.SetFocus: Exit Sub
  If lblHelp(4).Caption = "???" Then Beep: MsgBox "Concepto de Provisión C.T.S. no es valido; Verificar", vbExclamation: txtImportePvs.SetFocus: Exit Sub
  If txtRemunera(2).Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Percibida", vbExclamation: txtRemunera(2).SetFocus: Exit Sub
  If lblHelp(5).Caption = "???" Then Beep: MsgBox "Concepto de Remuneración percibida no es valido; Verificar", vbExclamation: txtRemunera(2).SetFocus: Exit Sub
  If txtCalculoCts(0).Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de calculo C.T.S. por dias", vbExclamation: txtCalculoCts(0).SetFocus: Exit Sub
  If lblHelp(6) = "???" Then Beep: MsgBox "Concepto de calculo C.T.S. por dias no es valido; Verificar", vbExclamation: txtCalculoCts(0).SetFocus: Exit Sub
  If txtCalculoCts(1).Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de calculo C.T.S. por meses", vbExclamation: txtCalculoCts(1).SetFocus: Exit Sub
  If lblHelp(7).Caption = "???" Then Beep: MsgBox "Concepto de calculo C.T.S. por meses no es valido; Verificar", vbExclamation: txtCalculoCts(1).SetFocus: Exit Sub
  If txtCalculoCts(2).Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de calculo C.T.S. por años", vbExclamation: txtCalculoCts(2).SetFocus: Exit Sub
  If lblHelp(8).Caption = "???" Then Beep: MsgBox "Concepto de calculo C.T.S. por años no es valido; Verificar", vbExclamation: txtCalculoCts(2).SetFocus: Exit Sub
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("pdoano", "remubasicacts", "remupromects", "remugraticts", "remutotalcts", "remunepvscts", "remupercibects", "remudiascts", "remumesescts", "remuanoscts", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, txtRemunera(0).Text, txtRemunera(1).Text, txtGratificacion.Text, txtImporteCts.Text, txtImportePvs.Text, txtRemunera(2).Text, txtCalculoCts(0).Text, txtCalculoCts(1).Text, txtCalculoCts(2).Text, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plparametroafp", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plparametroafp", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Ubico el registro ingresado o actualizado
  ShowScreen
  txtRemunera(0).SetFocus
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
  Me.Height = 5640: Me.Width = 6510
  Me.Left = 3580: Me.Top = 2500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Parametros de Certificado"
  lblTitle.Caption = "Parametros"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "edit": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "aceptar"
  aElemento(0, 2) = "Actualizar Información de " & lblTitle.Caption
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0, 1, 5     ' Concepto de remuneración ganada
    txtRemunera(IIf(n_IndexHelp = 5, 2, n_IndexHelp)).Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtRemunera(IIf(n_IndexHelp = 5, 2, n_IndexHelp)).SetFocus
   Case 2         ' Concepto de gratificacion
    txtGratificacion.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtGratificacion.SetFocus
   Case 3         ' Concepto de C.T.S.
    txtImporteCts.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtImporteCts.SetFocus
   Case 4         ' Concepto de provision
    txtImportePvs.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtImportePvs.SetFocus
   Case 6, 7, 8   ' Concepto de calculo de C.T.S.
    txtCalculoCts(n_IndexHelp - 6).Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtCalculoCts(n_IndexHelp - 6).SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0, 1, 5   ' Conceptos de remuneración ganada
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 2, 3, 4   ' Concepto de gratificacion y CTS
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 6, 7, 8 ' Conceptos de calculo de C.T.S.
    s_Registro = ps_ClsPlanilla & "F" & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
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
Private Sub txtCalculoCts_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCalculoCts(Index)
End Sub
Private Sub txtCalculoCts_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 6
End Sub
Private Sub txtCalculoCts_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCalculoCts_LostFocus(Index As Integer)
  lblHelp(Index + 6).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtCalculoCts(Index).Text, "CP")
End Sub
Private Sub txtGratificacion_GotFocus()
  gdl_Procedure.MarcaGet txtGratificacion
End Sub
Private Sub txtGratificacion_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 2
End Sub
Private Sub txtGratificacion_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtGratificacion_LostFocus()
  lblHelp(2).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtGratificacion.Text, "CP")
End Sub
Private Sub txtImporteCts_GotFocus()
  gdl_Procedure.MarcaGet txtImporteCts
End Sub
Private Sub txtImporteCts_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 3
End Sub
Private Sub txtImporteCts_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtImporteCts_LostFocus()
  lblHelp(3).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtImporteCts.Text, "CP")
End Sub
Private Sub txtImportePvs_GotFocus()
  gdl_Procedure.MarcaGet txtImporteCts
End Sub
Private Sub txtImportePvs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 4
End Sub
Private Sub txtImportePvs_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtImportePvs_LostFocus()
  lblHelp(4).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtImportePvs.Text, "CP")
End Sub
Private Sub txtRemunera_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtRemunera(Index)
End Sub
Private Sub txtRemunera_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtRemunera_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_LostFocus(Index As Integer)
  lblHelp(IIf(Index = 2, 5, Index)).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(Index).Text, "CP")
End Sub
