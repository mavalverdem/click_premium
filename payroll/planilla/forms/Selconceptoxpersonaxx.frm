VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fSelConceptoxPersona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "Selconceptoxpersona.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7740
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   4800
      TabIndex        =   22
      Top             =   480
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "."
      TabPicture(0)   =   "Selconceptoxpersona.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblconcepto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblHelp(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdHelp(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "desdemes"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "hastames"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Aceptar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Cancelar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "desdeano"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "hastaano"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtconcepto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox txtconcepto 
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
         Height          =   280
         Left            =   120
         MaxLength       =   8
         TabIndex        =   32
         Top             =   1560
         Width           =   675
      End
      Begin VB.ComboBox hastaano 
         Height          =   315
         Left            =   1920
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox desdeano 
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Cancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Aceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox hastames 
         Height          =   315
         Left            =   720
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox desdemes 
         Height          =   315
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Top             =   1560
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
         Index           =   1
         Left            =   1200
         TabIndex        =   34
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label lblconcepto 
         BackColor       =   &H80000000&
         Caption         =   "Concepto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc dcaRegistro 
      Height          =   330
      Left            =   45
      Top             =   5490
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Registro - 00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   5235
      Index           =   0
      Left            =   6960
      TabIndex        =   4
      Top             =   585
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9234
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
         Left            =   0
         TabIndex        =   14
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Registro"
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
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1635
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
         Picture         =   "Selconceptoxpersona.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2055
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
         Picture         =   "Selconceptoxpersona.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2760
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
         Picture         =   "Selconceptoxpersona.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3195
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
         Picture         =   "Selconceptoxpersona.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   4305
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
         Picture         =   "Selconceptoxpersona.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   13
         Tag             =   "0"
         Top             =   4740
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
         Picture         =   "Selconceptoxpersona.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1200
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
         Picture         =   "Selconceptoxpersona.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3615
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
         Picture         =   "Selconceptoxpersona.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   495
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
         Picture         =   "Selconceptoxpersona.frx":0108
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
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
      Begin VB.CommandButton CmdDescansosmed 
         Caption         =   "Enfermedad"
         Height          =   350
         Left            =   3400
         TabIndex        =   36
         Top             =   120
         Width           =   700
      End
      Begin VB.CheckBox checkAsistencia 
         Height          =   255
         Left            =   6060
         TabIndex        =   35
         Top             =   120
         Width           =   255
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   360
         Left            =   6315
         TabIndex        =   21
         Top             =   105
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         ButtonWidth     =   1879
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reportes"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A1"
                     Text            =   "Faltas"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A2"
                     Text            =   "por Concepto"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A3"
                     Text            =   "por Inasistencia"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.TextBox txtPeriodo 
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
         Height          =   280
         Left            =   1485
         MaxLength       =   8
         TabIndex        =   2
         Top             =   105
         Width           =   1150
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   2700
         TabIndex        =   19
         Top             =   105
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   5145
         TabIndex        =   17
         Top             =   75
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
         PictureUp       =   "Selconceptoxpersona.frx":0124
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   4740
         TabIndex        =   16
         Top             =   75
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
         PictureUp       =   "Selconceptoxpersona.frx":0140
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   5550
         TabIndex        =   18
         Top             =   75
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
         PictureUp       =   "Selconceptoxpersona.frx":015C
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo de Pago :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1320
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
         Left            =   3105
         TabIndex        =   3
         Top             =   150
         Width           =   195
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   15
      Top             =   585
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   8546
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2328"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2249"
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
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   0
      TabIndex        =   20
      Top             =   390
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
Attribute VB_Name = "fSelConceptoxPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(2)                               ' Array de inicio y fin de seleccion de registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_OptRegistro As String
Private cnn As ADODB.Connection
Private opcion As Integer

Public Sub GenerarRep_Proceso1()
Dim RstPersonas As ADODB.Recordset
Dim RstValores As ADODB.Recordset
Dim cadena As String
Dim ianno As Integer
Dim imes As Integer
Dim quemes As String
Dim cadenaperiodo As String
Dim Fila As Integer
Dim suma As Double
Dim sumatoria() As Double
Dim contar As Integer
'**********************
Dim s_FechaHora As String, s_OldMessage As String
Dim s_Representante As String, s_RegisPatronal As String
Dim s_Distrito  As String, s_FechaReporte As String
'**********************
Set RstPersonas = New ADODB.Recordset
Set RstValores = New ADODB.Recordset
'**********************
  
  contar = 0
  
'  If desdemes.Text = "" Or hastames = "" Then Exit Sub
'  If Int(desdeano.Text & Left(desdemes, 2)) > Int(hastaano.Text & Left(hastames, 2)) Then
'        MsgBox ("Rango no Valido")
'        Exit Sub
'  End If
 
  For ianno = Int(desdeano.Text) To Int(hastaano.Text)
            For imes = IIf(ianno = Int(desdeano.Text), Int(Left(desdemes.Text, 2)), 1) To IIf(ianno = Int(hastaano.Text), Int(Left(hastames.Text, 2)), 12)
                    Select Case imes
                        Case 1
                            quemes = "EN"
                        Case 2
                            quemes = "FE"
                        Case 3
                            quemes = "MA"
                        Case 4
                            quemes = "AB"
                        Case 5
                            quemes = "MY"
                        Case 6
                            quemes = "JU"
                        Case 7
                            quemes = "JL"
                        Case 8
                            quemes = "AG"
                        Case 9
                            quemes = "SE"
                        Case 10
                            quemes = "OC"
                        Case 11
                            quemes = "NO"
                        Case 12
                            quemes = "DI"
                    End Select
                    If opcion = 1 Then
                        cadenaperiodo = cadenaperiodo & quemes & Right(ianno, 2) & " "
                    Else
                        cadenaperiodo = cadenaperiodo & "    " & quemes & Right(ianno, 2) & " "
                    End If
                    contar = contar + 1
            Next
  Next
  
    If opcion = 2 Then
        If lblHelp(1) = "???" Or lblHelp(1) = "" Then Beep: MsgBox "Error en Concepto; Verificar", vbExclamation: txtconcepto.SetFocus: Exit Sub
    End If
  
    ReDim sumatoria(contar)

    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Representante = "": s_RegisPatronal = "": s_Distrito = "": s_FechaReporte = ""
    
    ' Cambio el Mensaje
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = IIf(opcion = 1, "DETALLE DE FALTAS", "DETALLE DE CONCEPTO")
    gdl_Procedure.ps_ReportName = IIf(opcion = 1, "rptfaltas", "rptimpconceptos")
    ReDim aElemento(3, 7): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    If opcion = 1 Then
        aElemento(2, 2) = "Periodo;" & " Desde " & Right(desdemes, Len(desdemes) - 4) & " del " & desdeano.Text & " Hasta " & Right(hastames, Len(hastames) - 4) & " del " & hastaano.Text & ";true"
    Else
        aElemento(2, 2) = "Periodo;" & " Desde " & Right(desdemes, Len(desdemes) - 4) & " del " & desdeano.Text & " Hasta " & Right(hastames, Len(hastames) - 4) & " del " & hastaano.Text & "Concepto : " & txtconcepto & " - " & lblHelp(1) & ";true"
    End If
    aElemento(2, 3) = "": aElemento(2, 4) = ""
    aElemento(2, 5) = "": aElemento(2, 6) = ""
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "codpsn varchar(100) NULL, apepaterno varchar(100) NULL, "
    s_Sql = s_Sql & "apematerno varchar(100) NULL, nombres varchar(100) NULL, "
    s_Sql = s_Sql & "periodo varchar(500) NULL, valores varchar(500) NULL ) "
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
             
    
    If SSTab1.Height = 1280 Then
    'PLASISTENCIA
    s_Sql = "SELECT asi.codpsn, psn.apepaterno, psn.apematerno, psn.nombres"
    s_Sql = s_Sql & " FROM plasistencia asi"
    s_Sql = s_Sql & " INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn"
    s_Sql = s_Sql & " WHERE left(asi.codpdo,4)>='" & Int(desdeano.Text & Left(desdemes, 2)) & "' and left(asi.codpdo,2)<='" & Int(hastaano.Text & Left(hastames, 2)) & "'"
    s_Sql = s_Sql & " AND asi.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & " AND asi.codpsn IN(SELECT valor FROM rangoimpresion"
    s_Sql = s_Sql & " WHERE proceso='" & s_OptRegistro & "'"
    s_Sql = s_Sql & " AND usrcre='" & ps_Usuario & "'"
    s_Sql = s_Sql & " AND fyhcre='" & s_FechaHora & "')"
    s_Sql = s_Sql & " GROUP BY asi.codpsn "
    s_Sql = s_Sql & " ORDER BY psn.apepaterno, psn.apematerno, psn.nombres "
    Else
    'PLRESULTADO
    s_Sql = "SELECT res.codpsn, psn.apepaterno, psn.apematerno, psn.nombres"
    s_Sql = s_Sql & " FROM plresultado res"
    s_Sql = s_Sql & " INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn"
    s_Sql = s_Sql & " WHERE concat(res.pdoano,res.pdomes)>='" & Int(desdeano.Text & Left(desdemes, 2)) & "' and concat(res.pdoano,res.pdomes)<='" & Int(hastaano.Text & Left(hastames, 2)) & "'"
    s_Sql = s_Sql & " AND res.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & " AND res.codpsn IN(SELECT valor FROM rangoimpresion"
    s_Sql = s_Sql & " WHERE proceso='" & s_OptRegistro & "'"
    s_Sql = s_Sql & " AND usrcre='" & ps_Usuario & "'"
    s_Sql = s_Sql & " AND fyhcre='" & s_FechaHora & "')"
    s_Sql = s_Sql & " GROUP BY res.codpsn "
    s_Sql = s_Sql & " ORDER BY psn.apepaterno, psn.apematerno, psn.nombres "
    End If
    
    
    RstPersonas.Open s_Sql, cnn, adOpenStatic, adLockOptimistic
    
    If RstPersonas.RecordCount = 0 Then GoTo Finalizar
      
    RstPersonas.MoveFirst
    For Fila = 0 To RstPersonas.RecordCount - 1

        cadena = ""
        suma = 0
        contar = 0
        For ianno = Int(desdeano.Text) To Int(hastaano.Text)
            For imes = IIf(ianno = Int(desdeano.Text), Int(Left(desdemes.Text, 2)), 1) To IIf(ianno = Int(hastaano.Text), Int(Left(hastames.Text, 2)), 12)
                          
                If opcion = 1 Then
                    s_Sql = " select sum(asi.diafalta) "
                    s_Sql = s_Sql & " from plasistencia asi"
                    s_Sql = s_Sql & " where left(asi.codpdo,4)='" & ianno & IIf(imes < 10, "0" & imes, imes) & "'"
                    s_Sql = s_Sql & " and asi.codcls='" & ps_ClsPlanilla & "'"
                    s_Sql = s_Sql & " and asi.codpsn='" & RstPersonas.Fields(0) & "'"
                Else
                    s_Sql = "SELECT sum(importe_mn) "
                    s_Sql = s_Sql & " from plresultado res"
                    s_Sql = s_Sql & " where res.pdoano='" & ianno & "' and res.pdomes='" & IIf(imes < 10, "0" & imes, imes) & "'"
                    s_Sql = s_Sql & " and res.codcls='" & ps_ClsPlanilla & "'"
                    s_Sql = s_Sql & " and res.codpsn='" & RstPersonas.Fields(0) & "' and res.codcpc='" & txtconcepto.Text & "'"
                End If
                
                RstValores.Open s_Sql, cnn, adOpenStatic, adLockOptimistic
                
                If RstValores.RecordCount > 0 Then
                    
                    If opcion = 1 Then
                        cadena = cadena & IIf(IsNull(RstValores.Fields(0)), "---0", IIf(Len(RstValores.Fields(0)) = 1, "---" & RstValores.Fields(0), IIf(Len(RstValores.Fields(0)) = 2, "--" & RstValores.Fields(0), "-" & RstValores.Fields(0)))) & " "
                        suma = suma + IIf(IsNull(RstValores.Fields(0)), 0, RstValores.Fields(0))
                        sumatoria(contar) = sumatoria(contar) + IIf(IsNull(RstValores.Fields(0)), 0, RstValores.Fields(0))
                    Else
                        cadena = cadena & IIf(IsNull(RstValores.Fields(0)), "----0.00", IIf(Len(Format(RstValores.Fields(0), "#.00")) = 4, "----" & Format(RstValores.Fields(0), "#.00"), IIf(Len(Format(RstValores.Fields(0), "#.00")) = 5, "---" & Format(RstValores.Fields(0), "#.00"), IIf(Len(Format(RstValores.Fields(0), "#.00")) = 6, "--" & Format(RstValores.Fields(0), "#.00"), IIf(Len(Format(RstValores.Fields(0), "#.00")) = 7, "-" & Format(RstValores.Fields(0), "#.00"), "" & Format(RstValores.Fields(0), "#.00")))))) & " "
                        suma = suma + IIf(IsNull(RstValores.Fields(0)), 0, Format(RstValores.Fields(0), "#.00"))
                        sumatoria(contar) = sumatoria(contar) + IIf(IsNull(RstValores.Fields(0)), 0, Format(RstValores.Fields(0), "#.00"))
                    End If
                Else
                    cadena = cadena & "----0.00" & " "
                End If
                
                contar = contar + 1
                RstValores.Close
            Next
        Next
    
        
        Select Case opcion
        Case 1
            If suma < 10 Then
                cadena = cadena & " ---" & suma
            ElseIf suma > 9 And suma < 100 Then
                cadena = cadena & " --" & suma
            ElseIf suma > 99 And suma < 1000 Then
                cadena = cadena & " -" & suma
            ElseIf suma > 999 And suma < 10000 Then
                cadena = cadena & " " & suma
            End If
        Case 2
            If suma = 0 Then
                cadena = cadena & " ----0.00"
            ElseIf suma > 0 And suma < 10 Then
                cadena = cadena & " ----" & Format(suma, "#.00")
            ElseIf suma > 9 And suma < 100 Then
                cadena = cadena & " ---" & Format(suma, "#.00")
            ElseIf suma > 99 And suma < 1000 Then
                cadena = cadena & " --" & Format(suma, "#.00")
            ElseIf suma > 999 And suma < 10000 Then
                cadena = cadena & " -" & Format(suma, "#.00")
            ElseIf suma > 9999 And suma < 100000 Then
                cadena = cadena & " " & Format(suma, "#.00")
            End If
        End Select
   
        s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " (codpsn,apepaterno,apematerno,nombres,periodo,valores) VALUES(' "
        s_Sql = s_Sql & RstPersonas.Fields(0) & "','" & RstPersonas.Fields(1) & "','" & RstPersonas.Fields(2) & "','" & RstPersonas.Fields(3) & "','" & cadenaperiodo & IIf(opcion = 1, " TOTAL", "    TOTAL") & "','" & cadena & "')"
        gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    RstPersonas.MoveNext
    Next
    
    cadena = ""
    suma = 0
    For contar = LBound(sumatoria) To UBound(sumatoria)
        suma = sumatoria(contar) + suma
            If contar = UBound(sumatoria) Then
                sumatoria(contar) = suma
                If opcion = 1 Then
                    cadena = cadena & " "
                Else
                    cadena = cadena & " "
                End If
            End If
        Select Case opcion
        Case 1
            If sumatoria(contar) < 10 Then
                cadena = cadena & "---" & Trim(sumatoria(contar)) & " "
            ElseIf sumatoria(contar) > 9 And sumatoria(contar) < 100 Then
                cadena = cadena & "--" & Trim(sumatoria(contar)) & " "
            ElseIf sumatoria(contar) > 99 And sumatoria(contar) < 1000 Then
                cadena = cadena & "-" & Trim(sumatoria(contar)) & " "
            ElseIf sumatoria(contar) > 999 And sumatoria(contar) < 10000 Then
                cadena = cadena & "" & Trim(sumatoria(contar)) & " "
            End If
        Case 2
            If sumatoria(contar) = 0 Then
                cadena = cadena & "----0.00"
            ElseIf sumatoria(contar) > 0 And sumatoria(contar) < 10 Then
                cadena = cadena & "----" & Format(sumatoria(contar), "#.00") & " "
            ElseIf suma > 9 And sumatoria(contar) < 100 Then
                cadena = cadena & "---" & Format(sumatoria(contar), "#.00") & " "
            ElseIf suma > 99 And sumatoria(contar) < 1000 Then
                cadena = cadena & "--" & Format(sumatoria(contar), "#.00") & " "
            ElseIf suma > 999 And sumatoria(contar) < 10000 Then
                cadena = cadena & "-" & Format(sumatoria(contar), "#.00") & " "
            ElseIf suma > 9999 And sumatoria(contar) < 100000 Then
                cadena = cadena & "" & Format(sumatoria(contar), "#.00") & " "
            End If
        End Select
    Next

    
    RstPersonas.Close
    
    s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " (codpsn,apepaterno,apematerno,nombres,periodo,valores) VALUES('->"
    s_Sql = s_Sql & "','TOTAL','','','" & cadenaperiodo & "','" & cadena & "')"
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = " SELECT codpsn, apepaterno, apematerno, nombres,"
    s_Sql = s_Sql & " periodo, valores "
    s_Sql = s_Sql & " FROM tmp" & gdl_Procedure.ps_ReportName & " order by codpsn "
            
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (0), False, True, False, True, True, aElemento, aElementos, porstRecordset

Finalizar:
    Set porstRecordset = Nothing
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    ' Elimino la tabla temporal
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
    SSTab1.Visible = False
    Toolbar1.Enabled = True
End Sub


' Instancia del formulario activo
'[
Private Sub RecuperaRegistros(ByVal s_Orden As String)
  
  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, codpsn, apepaterno, apematerno, nombres,"
  s_Sql = s_Sql & " fecnacimiento, ubigeonac, naciextrapsn, sexopsn,"
  s_Sql = s_Sql & " refedirec, codvia, nomviadirec, numerdirec,"
  s_Sql = s_Sql & " intedirec, codzona, nomzondirec, ubigeodir,"
  s_Sql = s_Sql & " estcivilpsn, numhijo, numdepen, coddci, numdociden,"
  s_Sql = s_Sql & " numdocmil, telefono, celular, dctojudicial, pordsctojudi, fotopsn,"
  s_Sql = s_Sql & " fecingreso, codtpt, codcgo, cgoconfianza, codpfs, jornadalaboral,"
  s_Sql = s_Sql & " codcco, codafp, numeroafp, pagodolar, codbcopago,"
  s_Sql = s_Sql & " cuentapago, ctsdolar, codbcocts, cuentacts, codeps,"
  s_Sql = s_Sql & " regpension, fecingregpen, essvida, cobsctr, afilsindical,"
  s_Sql = s_Sql & " remintegralgrati, remuneta, netocpc, variacpc, imporemuneto,"
  s_Sql = s_Sql & " fecbaja, estadopsn"
  s_Sql = s_Sql & " FROM plpersonal"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & " AND estadopsn" & IIf(ribParametro(1).Value, "<>'I'", "='I'")
  End If
  s_Sql = s_Sql & " ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly
  
  ' Inicializo los rangos de impresion
  as_SelRegistro(0) = "": as_SelRegistro(1) = ""
  If dcaRegistro.Recordset.RecordCount > 0 Then
    dcaRegistro.Recordset.MoveLast: as_SelRegistro(1) = dcaRegistro.Recordset.Bookmark
    dcaRegistro.Recordset.MoveFirst: as_SelRegistro(0) = dcaRegistro.Recordset.Bookmark
  End If

End Sub
Private Sub Aceptar_Click()

Dim s_FechaHora As String
Dim s_OldMessage As String

If desdemes.Text = "" Or hastames = "" Then Exit Sub
If Int(desdeano.Text & Left(desdemes, 2)) > Int(hastaano.Text & Left(hastames, 2)) Then
        MsgBox ("Rango no Valido")
        Exit Sub
End If

If opcion = 1 Or opcion = 2 Then
  Call GenerarRep_Proceso1
End If
MsgBox desdeano.Text & Left(desdemes, 2)
MsgBox hastaano.Text & Left(hastames, 2)
  
'Mayo 2015
'Proceso para generacion de reporte de Inasistencias (Enfermedad, Natalidad y Accidente)
If opcion = 3 Then
    
     s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    'gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), 1, s_TitleTable
    
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    ' Cambio el Mensaje
    's_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
     tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
     gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "ANALISIS DE INASISTENCIAS"
    gdl_Procedure.ps_ReportName = "cstinasistenciasdet"
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    'aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    'aElemento(2, 2) = "Periodo;" & " Desde " & Date & " Hasta " & Date & ";true"
        aElemento(2, 2) = "Periodo;" & " Desde " & Right(desdemes, Len(desdemes) - 4) & " del " & desdeano.Text & " Hasta " & Right(hastames, Len(hastames) - 4) & " del " & hastaano.Text & ";true"
    
    'codmdi_enfer
    ' [ Generación e impresión de información para el reporte
        
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "codpdo varchar(12) NULL, codpsn varchar(10) NULL, "
    s_Sql = s_Sql & "apellidosnombre varchar(80) NULL, dias_enfermedad NUMERIC NULL, "
    s_Sql = s_Sql & "dias_natalidad NUMERIC NULL,dias_accidente NUMERIC NULL, "
    s_Sql = s_Sql & "cod_motivoina varchar(2) NULL,des_motivoina VARCHAR(80), "
    s_Sql = s_Sql & "fechaini_ina date NULL, fechafin_ina date NULL, fecingreso date NULL, fecbaja date NULL ) "
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "select CONCAT(LEFT(codpdo,4),'-',RIGHT(codpdo,2)) AS codpdo,asis.codpsn, CONCAT(apematerno,' ', apematerno, ' ', nombres) as apellidosnombres, "
    s_Sql = s_Sql & "0 as dias_enfermedad,asis.diaprepostnatal as dias_natalidad,0 as dias_accidente, tsus.codtsu as cod_motivoina, "
    s_Sql = s_Sql & "destsu As des_motivoina, fechaini_natal As fechaini_ina, fechafin_natal As fechafin_ina, per.fecingreso, per.fecbaja "
    s_Sql = s_Sql & "FROM plasistencia asis "
    s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_natal=tsus.codtsu "
    s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn = per.codpsn "
    s_Sql = s_Sql & "AND asis.codcls='" & ps_ClsPlanilla & "' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
    's_Sql = s_Sql & "AND RTRIM(left(asis.codpdo,4))='2015' AND LTRIM(right(asis.codpdo,2)) BETWEEN '01' AND '04' "
    s_Sql = s_Sql & "AND asis.codpdo>='" & desdeano.Text & Left(desdemes, 2) & "' and asis.codpdo<= '" & hastaano.Text & Left(hastames, 2) & "' "
    
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "select  CONCAT(LEFT(codpdo,4),'-',RIGHT(codpdo,2)) AS codpdo, asis.codpsn,  CONCAT(per.apepaterno,' ', per.apematerno, ' ', per.nombres) as apellidosnombres, "
    s_Sql = s_Sql & "asis.enfermedad as dias_enfermedad,0 as dias_natalidad, 0 as dias_accidente,tsus.codtsu as cod_motivoina, "
    s_Sql = s_Sql & "tsus.destsu as des_motivoina, fechaini_enfer as fechaini_ina, fechafin_enfer as fechafin_ina, per.fecingreso, per.fecbaja "
    s_Sql = s_Sql & "FROM plasistencia asis "
    s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_enfer=tsus.codtsu "
    s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn = per.codpsn "
    s_Sql = s_Sql & "AND asis.codcls='" & ps_ClsPlanilla & "' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
    s_Sql = s_Sql & "AND asis.codpdo>='" & desdeano.Text & Left(desdemes, 2) & "' and asis.codpdo<= '" & hastaano.Text & Left(hastames, 2) & "' "
    
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT  CONCAT(LEFT(codpdo,4),'-',RIGHT(codpdo,2)) AS codpdo,asis.codpsn,  CONCAT(apematerno,' ', apematerno, ' ', nombres) as apellidosnombres, "
    s_Sql = s_Sql & "0 as dias_enfermedad, 0 as dias_natalidad, asis.accidente as dias_accidente, tsus.codtsu as cod_motivoina, "
    s_Sql = s_Sql & "tsus.destsu as des_motivoina, fechaini_accid as fechaini_ina, fechafin_accid as fechafin_ina, per.fecingreso, per.fecbaja "
    s_Sql = s_Sql & "FROM plasistencia asis "
    s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_accid=tsus.codtsu "
    s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn = per.codpsn "
     
    s_Sql = s_Sql & "AND asis.codcls='" & ps_ClsPlanilla & "' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
    s_Sql = s_Sql & "AND asis.codpdo>='" & desdeano.Text & Left(desdemes, 2) & "' and asis.codpdo<= '" & hastaano.Text & Left(hastames, 2) & "' "
    s_Sql = s_Sql & "order by codpsn, codpdo,fechaini_ina;"
    
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    'Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (0), False, True, False, True, True, aElemento, aElementos, porstRecordset
    
    
    If porstRecordset.RecordCount = 0 Then GoTo Finalizar
    MuestraMensaje s_OldMessage
End If
Finalizar:
  Set porstRecordset = Nothing
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  'Elimino la tabla temporal
  s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql

End Sub

Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_OldMessage As String
  Dim s_Representante As String, s_RegisPatronal As String
  Dim s_Distrito  As String, s_FechaReporte As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Visualizar o analizar registro
    If txtPeriodo.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0).Caption = "" Or lblHelp(0).Caption = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    Me.Tag = s_MdoData_Vis
    If s_OptRegistro = "exepcional" Then
      fRemunerExcepcional.Show
    ElseIf s_OptRegistro = "asistencia" Then
      If tdbRegistro.SelBookmarks.Count > 0 Then If MsgBox("¿ Estás Seguro de actualizar el personal seleccionado ?", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
      fAbcAsistencia.Show vbModal
    ElseIf s_OptRegistro = "consulxcpc" Then
      fMenu.Tag = s_OptRegistro
      o_Consultaxcpc.Show
    End If
   Case 1, 2  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 3 ' Busqueda de registro
    Set go_tdbBusqueda = tdbRegistro
    Set go_dcaBusqueda = dcaRegistro
    gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
    fBusqueda.Show vbModal
   Case 4, 5, 6 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), (Index - 4), s_TitleTable
   Case 7, 8  ' Opciones de impresión
    If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Representante = "": s_RegisPatronal = "": s_Distrito = "": s_FechaReporte = ""
    
    ' Cambio el Mensaje
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = IIf(s_OptRegistro = "exepcional", "MOVIMIENTOS EXEPCIONALES", IIf(s_OptRegistro = "asistencia", "DETALLE DE ASISTENCIA Y PUNTUALIDAD", "RESULTADOS DE PROCESO DE CÁLCULO"))
    gdl_Procedure.ps_ReportName = IIf(s_OptRegistro = "exepcional", "rptremexepci", IIf(s_OptRegistro = "asistencia", "rptasistencia", "cstconcextraba"))
    ReDim aElemento(3, 7): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Periodo;" & Trim(txtPeriodo.Text) & " - " & Trim(lblHelp(0).Caption) & ";true"
    aElemento(2, 3) = "": aElemento(2, 4) = ""
    aElemento(2, 5) = "": aElemento(2, 6) = ""
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    If s_OptRegistro = "exepcional" Then
      s_Sql = "SELECT psn.codpsn, psn.apepaterno, psn.apematerno, psn.nombres,"
      s_Sql = s_Sql & " rex.codpdo, rex.codcpc, rex.codmon, cpc.descpc, cpc.tipocpc, rex.imporemune"
      s_Sql = s_Sql & " FROM plpersonal psn"
      s_Sql = s_Sql & " INNER JOIN plremuexce rex ON psn.codcls=rex.codcls AND psn.codpsn=rex.codpsn"
      s_Sql = s_Sql & " INNER JOIN plconcepto cpc ON rex.codcpc=cpc.codcpc"
      s_Sql = s_Sql & " WHERE rex.codpdo='" & Trim(txtPeriodo.Text) & "'"
    ElseIf s_OptRegistro = "asistencia" Then
      If checkAsistencia.Value = Checked Then
        s_Sql = "SELECT psn.codpsn, psn.apepaterno, psn.apematerno, concat(psn.nombres,' ', asi.codpdo),"
      Else
        s_Sql = "SELECT psn.codpsn, psn.apepaterno, psn.apematerno, psn.nombres,"
      End If
      s_Sql = s_Sql & " asi.codpdo, asi.diatrabajo, asi.horanormal, asi.horatipo1,"
      s_Sql = s_Sql & " asi.horatipo2, asi.horatipo3, asi.diafalta, asi.tardanza,"
      s_Sql = s_Sql & " asi.diaprepostnatal, asi.accidente, asi.diavacaciones,"
      s_Sql = s_Sql & " asi.enfermedad, asi.licencia, asi.permisos, asi.fechainivacacion,"
      s_Sql = s_Sql & " asi.fechafinvacacion, asi.pdovaca1, asi.fechainivaca1, asi.fechafinvaca1,"
      s_Sql = s_Sql & " asi.pdovaca2, asi.fechainivaca2, asi.fechafinvaca2, asi.dialiquidacion,"
      s_Sql = s_Sql & " asi.liquidavacacion, asi.diagratificacion, asi.fechacese, asi.fechainiliqvaca,"
      s_Sql = s_Sql & " asi.fechafinliqvaca, asi.tercerturno, asi.diasuspension, asi.opcional "
      s_Sql = s_Sql & " FROM plpersonal psn"
      s_Sql = s_Sql & " INNER JOIN plasistencia asi ON psn.codcls=asi.codcls AND psn.codpsn=asi.codpsn"
      If checkAsistencia.Value = Checked Then
        s_Sql = s_Sql & " WHERE LEFT(asi.codpdo,4)='" & ps_Ano & "' "
      Else
        s_Sql = s_Sql & " WHERE asi.codpdo='" & Trim(txtPeriodo.Text) & "'"
      End If
    ElseIf s_OptRegistro = "consulxcpc" Then
      s_Sql = "SELECT res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS apellidosnombres,"
      s_Sql = s_Sql & " res.secuencia, res.codcpc, cpc.descpc , res.tipocpc, cxp.defaultcpc, cxp.clasecpc, res.impbolecpc, res.importe_mn, res.importe_me"
      s_Sql = s_Sql & " FROM plpersonal psn"
      s_Sql = s_Sql & " INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn"
      s_Sql = s_Sql & " INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc"
      s_Sql = s_Sql & " INNER JOIN plconceplanilla cxp ON res.codcls=cxp.codcls AND res.codcpc=cxp.codcpc"
      s_Sql = s_Sql & " WHERE res.codpdo='" & Trim(txtPeriodo.Text) & "'"
    End If
    s_Sql = s_Sql & " AND psn.codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & " AND psn.codpsn IN(SELECT valor FROM rangoimpresion"
    s_Sql = s_Sql & " WHERE proceso='" & s_OptRegistro & "'"
    s_Sql = s_Sql & " AND usrcre='" & ps_Usuario & "'"
    s_Sql = s_Sql & " AND fyhcre='" & s_FechaHora & "')"
    s_Sql = s_Sql & " ORDER BY " & aElemento(0, 1) & IIf(s_OptRegistro = "consulxcpc", ", secuencia, codcpc", "")
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 7), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    ' Elimino la tabla temporal
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
  End Select

End Sub


Private Sub CmdDescansosmed_Click()
Dim s_FechaHora As String
Dim s_OldMessage As String


 s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
'gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), 1, s_TitleTable

' Verifico que existan registros seleccionados
If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
' Cambio el Mensaje
's_OldMessage = fMenu.panMessage.Caption
MuestraMensaje "Procesando Información ..."
' Barro el arreglo de registros marcadas (bookmarks)
For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
 tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
 gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
Next n_Index

' Parametros de Impresión
gdl_Procedure.ps_ReportTitle = "ANALISIS DE INASISTENCIAS"
gdl_Procedure.ps_ReportName = "cstinasistenciasdet"
ReDim aElemento(3, 3): ReDim aElementos(2)
' Parametros del Reporte
aElemento(0, 0) = ps_CodEmpresa
aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
aElemento(0, 2) = ""
' Formulas del Reporte
'aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
' Parametros de campos del Reporte
aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
aElemento(2, 2) = "Periodo;" & " Desde " & Date & " Hasta " & Date & ";true"


'codmdi_enfer
' [ Generación e impresión de información para el reporte
    
s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  
s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
s_Sql = s_Sql & "codpdo varchar(12) NULL, codpsn varchar(10) NULL, "
s_Sql = s_Sql & "apellidosnombre varchar(80) NULL, dias_enfermedad NUMERIC NULL, "
s_Sql = s_Sql & "dias_natalidad NUMERIC NULL,dias_accidente NUMERIC NULL, "
s_Sql = s_Sql & "cod_motivoina varchar(2) NULL,des_motivoina VARCHAR(80), "
s_Sql = s_Sql & "fechaini_ina date NULL, fechafin_ina date NULL, fecingreso date NULL, fecbaja date NULL ) "
gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql

s_Sql = "select CONCAT(LEFT(codpdo,4),'-',RIGHT(codpdo,2)) AS codpdo,asis.codpsn, CONCAT(apematerno,' ', apematerno, ' ', nombres) as apellidosnombres, "
s_Sql = s_Sql & "0 as dias_enfermedad,asis.diaprepostnatal as dias_natalidad,0 as dias_accidente, tsus.codtsu as cod_motivoina, "
s_Sql = s_Sql & "destsu As des_motivoina, fechaini_natal As fechaini_ina, fechafin_natal As fechafin_ina, per.fecingreso, per.fecbaja "
s_Sql = s_Sql & "FROM plasistencia asis "
s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_natal=tsus.codtsu "
s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn = per.codpsn "
s_Sql = s_Sql & "AND asis.codcls='01' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
s_Sql = s_Sql & "AND RTRIM(left(asis.codpdo,4))='2015' AND LTRIM(right(asis.codpdo,2)) BETWEEN '01' AND '04' "

s_Sql = s_Sql & "UNION "
s_Sql = s_Sql & "select  CONCAT(LEFT(codpdo,4),'-',RIGHT(codpdo,2)) AS codpdo, asis.codpsn,  CONCAT(per.apepaterno,' ', per.apematerno, ' ', per.nombres) as apellidosnombres, "
s_Sql = s_Sql & "asis.enfermedad as dias_enfermedad,0 as dias_natalidad, 0 as dias_accidente,tsus.codtsu as cod_motivoina, "
s_Sql = s_Sql & "tsus.destsu as des_motivoina, fechaini_enfer as fechaini_ina, fechafin_enfer as fechafin_ina, per.fecingreso, per.fecbaja "
s_Sql = s_Sql & "FROM plasistencia asis "
s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_enfer=tsus.codtsu "
s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn = per.codpsn "
s_Sql = s_Sql & "AND asis.codcls='01' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
s_Sql = s_Sql & "AND RTRIM(left(asis.codpdo,4))='2015' AND LTRIM(right(asis.codpdo,2)) BETWEEN '01' AND '04' "

s_Sql = s_Sql & "UNION "
s_Sql = s_Sql & "SELECT  CONCAT(LEFT(codpdo,4),'-',RIGHT(codpdo,2)) AS codpdo,asis.codpsn,  CONCAT(apematerno,' ', apematerno, ' ', nombres) as apellidosnombres, "
s_Sql = s_Sql & "0 as dias_enfermedad, 0 as dias_natalidad, asis.accidente as dias_accidente, tsus.codtsu as cod_motivoina, "
s_Sql = s_Sql & "tsus.destsu as des_motivoina, fechaini_accid as fechaini_ina, fechafin_accid as fechafin_ina, per.fecingreso, per.fecbaja "
s_Sql = s_Sql & "FROM plasistencia asis "
s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_accid=tsus.codtsu "
s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn = per.codpsn "
 
s_Sql = s_Sql & "AND asis.codcls='01' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
s_Sql = s_Sql & "AND RTRIM(left(asis.codpdo,4))='2015' AND LTRIM(right(asis.codpdo,2)) BETWEEN '01' AND '04' "
s_Sql = s_Sql & "order by codpsn, codpdo;"

Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
'Ejecuto reporte y saco de memoria la información
gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (0), False, True, False, True, True, aElemento, aElementos, porstRecordset


If porstRecordset.RecordCount = 0 Then GoTo Finalizar
MuestraMensaje s_OldMessage

Finalizar:
  Set porstRecordset = Nothing
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  'Elimino la tabla temporal
  s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
End Sub

Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Periodo de Pago
    gdl_Procedure.DefineStyleGrilla tdbHelp, "Periodo de Pago", 2
    tdbHelp.Columns(0).DataField = "codpdo": tdbHelp.Columns(1).DataField = "despdo"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ped", "codpdo", IIf(s_OptRegistro = "consulxcpc", s_Estado_Ina, s_Estado_Blq) & ps_ClsPlanilla & ps_Ano, "")
   Case 1
    gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos", 2
   tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
   s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", ps_ClsPlanilla & "F" & "0", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = panToolBar(1).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = panToolBar(1).Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdResumDesmedco_Click()
Dim s_FechaHora As String
Dim s_OldMessage As String
Dim n_DiasSubsidio As Integer

n_DiasSubsidio = 6
s_FechaHora = Format(Now, s_FmtFeHoMysql_0)

' Verifico que existan registros seleccionados
If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
' Cambio el Mensaje
's_OldMessage = fMenu.panMessage.Caption
MuestraMensaje "Procesando Información ..."
' Barro el arreglo de registros marcadas (bookmarks)
For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
 tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
 gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
Next n_Index

' Parametros de Impresión
gdl_Procedure.ps_ReportTitle = "CONSOLIDADO DE DESCANSOS MEDICOS CON SUBSIDIO"
gdl_Procedure.ps_ReportName = "cstconsoldesmdico"
ReDim aElemento(3, 3): ReDim aElementos(2)
' Parametros del Reporte
aElemento(0, 0) = ps_CodEmpresa
aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
aElemento(0, 2) = ""
' Formulas del Reporte
'aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
' Parametros de campos del Reporte
aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
aElemento(2, 2) = "Periodo;" & " Desde " & Date & " Hasta " & Date & ";true"


' [ Generación e impresión de información para el reporte
    
s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  
s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
s_Sql = s_Sql & "codcls varchar(2) NULL, aniopdo varchar(4) NULL, "
s_Sql = s_Sql & "codpsn varchar(11) NULL, diasenfer numeric NULL, "
s_Sql = s_Sql & "apellidosnombre varchar(80) NULL, "
s_Sql = s_Sql & "destsu varchar(80) NULL, fecingreso date NULL,detcco varchar(40) ) "
gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql


s_Sql = "SELECT per.codcls,RTRIM(left(codpdo,4)) AS aniopdo,asis.codpsn, SUM(enfermedad) as diasenfer,concat(nombres, ' ', "
s_Sql = s_Sql & "apepaterno) as nompersonal, DATE_FORMAT(fecingreso,'%d/%m/%Y') as fecingreso,coc.detcco "
s_Sql = s_Sql & "FROM plasistencia asis "
s_Sql = s_Sql & "INNER JOIN pltipsusp tsus ON asis.codmdi_enfer=tsus.codtsu "
s_Sql = s_Sql & "INNER JOIN plpersonal per ON asis.codpsn=per.codpsn "
s_Sql = s_Sql & "INNER JOIN cocco coc ON per.codcco=coc.CodCCo "
's_Sql = s_Sql & "AND asis.codcls='01' AND asis.codpsn IN ('0000033','0000072','0000086','0000087')"

s_Sql = s_Sql & "AND asis.codcls='01' AND asis.codpsn IN (SELECT valor FROM rangoimpresion WHERE proceso='" & s_OptRegistro & "' "
s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' AND fyhcre='" & s_FechaHora & "') "
s_Sql = s_Sql & "AND RTRIM(left(codpdo,4))=2014 AND LTRIM(right(codpdo,2)) BETWEEN '01' AND '12' "
s_Sql = s_Sql & "GROUP BY asis.codpsn Having Sum(enfermedad) > 6 "



Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
'Ejecuto reporte y saco de memoria la información
gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (0), False, True, False, True, True, aElemento, aElementos, porstRecordset


If porstRecordset.RecordCount = 0 Then GoTo Finalizar
MuestraMensaje s_OldMessage

Finalizar:
  Set porstRecordset = Nothing
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  'Elimino la tabla temporal
  s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
End Sub





Private Sub dcaRegistro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  If s_OptRegistro = "exepcional" Then
    If FormVisible("fRemunerExcepcional") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fRemunerExcepcional.RecuperaRegistros "codcpc"
      End If
    End If
  ElseIf s_OptRegistro = "asistencia" Then
    If FormVisible("fAbcAsistencia") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fAbcAsistencia.ShowScreen
      End If
    End If
  ElseIf s_OptRegistro = "consulxcpc" Then
    If FormVisible("fConsultaCalculo") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        o_Consultaxcpc.RecuperaRegistros
      End If
    End If
  End If

End Sub
Private Sub Form_Activate()
  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim Item As New ValueItem
  
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
  cnn.CursorLocation = adUseClient
  cnn.Open
  
  SSTab1.Visible = False
  
  For n_Index = (Val(ps_Ano) - 5) To (Val(ps_Ano) + 5)
      desdeano.AddItem n_Index
  Next n_Index
  desdeano.ListIndex = 5
  
  For n_Index = (Val(ps_Ano) - 5) To (Val(ps_Ano) + 5)
      hastaano.AddItem n_Index
  Next n_Index
  hastaano.ListIndex = 5
  
  For n_Index = 1 To 12: desdemes.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  For n_Index = 1 To 12: hastames.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  
  ' Establece posición del formulario
  Me.Height = 6340: Me.Width = 7830
  Me.Left = 105: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro

  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Trabajador(es)"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Apellido Paterno", "Apellido Materno", "Nombre(s)", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "apepaterno", "apematerno", "nombres", "estadopsn")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1080, 1616.33, 1616.33, 1616.33, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0)
  Next n_Index
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(4).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(4).ValueItems.Translate = True
  For n_Index = 0 To 5
    tdbRegistro.Columns(4).ValueItems.Add Item
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "A", "V", "L", "P", "O", "I")
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadovo", "estadnok", "estadopk", "estadopn", "procenok") & ".bmp")
  Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbRegistro.AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(9, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "registro": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "seleccio", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Selecciona y Edita Registro", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codbco", "desbco")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 934.7402, 3255.071)
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Periodo de Pago", 2
  ']
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  
End Sub
Private Sub Form_Unload(Cancel As Integer)

  If s_OptRegistro = "exepcional" Then
    If FormVisible("fRemunerExcepcional") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fRemunerExcepcional.Caption, vbExclamation
      Cancel = True
      Exit Sub
    End If
  ElseIf s_OptRegistro = "asistencia" Then
    If FormVisible("fAbcAsistencia") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fAbcAsistencia.Caption, vbExclamation
      Cancel = True
      Exit Sub
    End If
  ElseIf s_OptRegistro = "consulxcpc" Then
    If FormVisible("fConsultaCalculo") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & o_Consultaxcpc.Caption, vbExclamation
      Cancel = True
      Exit Sub
    End If
  End If
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
  ' Habilito la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = True

End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index
  
End Sub
Private Sub ribParametro_Click(Index As Integer, Value As Integer)
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Periodo de pago
    txtPeriodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo.SetFocus
   Case 1       ' Conceptos
    txtconcepto = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtconcepto.SetFocus
    
  End Select
   
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Periodo de Pago
    If s_OptRegistro = "consulxcpc" Or s_OptRegistro = "anrenta5ta" Then
      s_Sql = gdl_Funcion.HelpTablas("ped", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & ps_Ano, "")
    Else
      s_Sql = gdl_Funcion.HelpTablas("ped", tdbHelp.Columns(ColIndex).DataField, s_Estado_Blq & ps_ClsPlanilla & ps_Ano, "")
    End If
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, "PR")
End Sub
Private Sub txtConcepto_GotFocus()
  gdl_Procedure.MarcaGet txtconcepto
End Sub
Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtConcepto_LostFocus()
  lblHelp(1).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtconcepto, "CP")
End Sub

Private Sub toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
desdeano.ListIndex = 5
hastaano.ListIndex = 5
desdemes.Text = ""
hastames.Text = ""
txtconcepto.Text = ""
lblHelp(1).Caption = ""
Toolbar1.Enabled = False
Select Case ButtonMenu.Key
Case "A1"
    SSTab1.Height = 1280
    SSTab1.Visible = True
    opcion = 1
Case "A2"
    SSTab1.Height = 1950
    SSTab1.Visible = True
    opcion = 2
Case "A3"
    SSTab1.Height = 1950
    SSTab1.Visible = True
    opcion = 3
End Select


End Sub
Private Sub Cancelar_Click()
    SSTab1.Visible = False
    Toolbar1.Enabled = True
End Sub
