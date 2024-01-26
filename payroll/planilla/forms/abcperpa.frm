VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fAbcPeriodoPago 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4740
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "abcperpa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   7740
   Begin TabDlg.SSTab tabRegister 
      Height          =   3540
      Left            =   75
      TabIndex        =   34
      Top             =   600
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6244
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
      TabPicture(0)   =   "abcperpa.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCuadro(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescripcion"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbTipo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "abcperpa.frx":0028
         Left            =   4650
         List            =   "abcperpa.frx":003B
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2565
         Width           =   1800
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         MaxLength       =   50
         TabIndex        =   3
         Top             =   615
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
         Top             =   270
         Width           =   1185
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   915
         Index           =   1
         Left            =   4650
         TabIndex        =   15
         Top             =   1005
         Width           =   1950
         _Version        =   65536
         _ExtentX        =   3440
         _ExtentY        =   1614
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
            TabIndex        =   16
            Top             =   285
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&No Procesado"
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
            TabIndex        =   17
            Top             =   585
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Procesado"
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
         Height          =   1875
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   1005
         Width           =   4350
         _Version        =   65536
         _ExtentX        =   7673
         _ExtentY        =   3307
         _StockProps     =   14
         Caption         =   " Fechas"
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
         Begin VB.TextBox txtFecha 
            Height          =   300
            Left            =   300
            MaxLength       =   4
            TabIndex        =   12
            Top             =   1365
            Width           =   975
         End
         Begin VB.ComboBox cmbPeriodo 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "abcperpa.frx":0076
            Left            =   2430
            List            =   "abcperpa.frx":0078
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1365
            Width           =   1590
         End
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   615
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            Format          =   137756673
            CurrentDate     =   37515
         End
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   1
            Left            =   1530
            TabIndex        =   8
            Top             =   615
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   137756673
            CurrentDate     =   37515
         End
         Begin MSComCtl2.DTPicker dtpFechas 
            Height          =   300
            Index           =   2
            Left            =   2910
            TabIndex        =   10
            Top             =   615
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            CalendarForeColor=   12582912
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   128
            Format          =   137756673
            CurrentDate     =   37515
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Año :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   11
            Top             =   1065
            Width           =   1005
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Mes :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   2445
            TabIndex        =   13
            Top             =   1065
            Width           =   1005
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   750
            Index           =   0
            Left            =   105
            Shape           =   4  'Rounded Rectangle
            Top             =   1020
            Width           =   4140
         End
         Begin VB.Label lblDato 
            Caption         =   "Pago :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   2910
            TabIndex        =   9
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Hasta :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   1530
            TabIndex        =   7
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Desde :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   330
            Width           =   1005
         End
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "yyyymm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   8
         Left            =   2640
         TabIndex        =   35
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblDato 
         Caption         =   "Tipo Periodo :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   4650
         TabIndex        =   18
         Top             =   2265
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   660
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
         Top             =   315
         Width           =   1000
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   20
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
         TabIndex        =   21
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
         Picture         =   "abcperpa.frx":007A
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6300
         TabIndex        =   22
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
         Picture         =   "abcperpa.frx":0096
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
         TabIndex        =   23
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   24
      Top             =   4230
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
         TabIndex        =   25
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
         Picture         =   "abcperpa.frx":00B2
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4545
         TabIndex        =   26
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
         Picture         =   "abcperpa.frx":00CE
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2835
         TabIndex        =   27
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
         Picture         =   "abcperpa.frx":00EA
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2445
         TabIndex        =   28
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
         Picture         =   "abcperpa.frx":0106
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   3540
      Index           =   0
      Left            =   6960
      TabIndex        =   29
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6244
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         Picture         =   "abcperpa.frx":0122
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   32
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
         Picture         =   "abcperpa.frx":013E
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   33
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
         Picture         =   "abcperpa.frx":015A
      End
   End
End
Attribute VB_Name = "fAbcPeriodoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private s_Periodo As String                             ' Codigo del registro
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
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fPeriodoPago.dcaRegistro.Recordset!codpdo.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fPeriodoPago.dcaRegistro.Recordset!despdo.DefinedSize
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(2), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txtFecha, Year(Date), Me.Tag, False, fPeriodoPago.dcaRegistro.Recordset!anopdo.DefinedSize
    n_Index = Month(Date) - 1
    gdl_Procedure.EditCombo "AT", cmbPeriodo, n_Index, Me.Tag, False
    gdl_Procedure.EditCombo "AT", cmbtipo, -1, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, False
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fPeriodoPago.dcaRegistro.Recordset!codpdo, Me.Tag, True, fPeriodoPago.dcaRegistro.Recordset!codpdo.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fPeriodoPago.dcaRegistro.Recordset!despdo), Me.Tag, False, fPeriodoPago.dcaRegistro.Recordset!despdo.DefinedSize
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), fPeriodoPago.dcaRegistro.Recordset!fechaini, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), fPeriodoPago.dcaRegistro.Recordset!fechafin, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(2), fPeriodoPago.dcaRegistro.Recordset!fechapago, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditText "AT", txtFecha, fPeriodoPago.dcaRegistro.Recordset!anopdo, Me.Tag, False, fPeriodoPago.dcaRegistro.Recordset!anopdo.DefinedSize
    n_Index = (fPeriodoPago.dcaRegistro.Recordset!mespdo)
    gdl_Procedure.EditCombo "AT", cmbPeriodo, (n_Index - 1), Me.Tag, False
    n_Index = IIf(fPeriodoPago.dcaRegistro.Recordset!tpopdo = "N", 0, IIf(fPeriodoPago.dcaRegistro.Recordset!tpopdo = "G", 1, IIf(fPeriodoPago.dcaRegistro.Recordset!tpopdo = "V", 2, IIf(fPeriodoPago.dcaRegistro.Recordset!tpopdo = "L", 3, 4))))
    gdl_Procedure.EditCombo "AT", cmbtipo, n_Index, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fPeriodoPago.dcaRegistro.Recordset!estadopdo = s_Estado_Ina), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fPeriodoPago.dcaRegistro.Recordset!estadopdo <> s_Estado_Ina), Me.Tag, False
  End If

End Sub
Private Sub cmdAction_Click(Index As Integer)

  ' Valido que el peiodo no se encuentre procesado
  If optEstado(1).Value And Index <> 0 Then Beep: MsgBox "Periodo No se puede Actualizar se encuentra Procesado", vbExclamation: Me.Tag = s_MdoData_Vis: Exit Sub
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
    s_Periodo = Trim$(txtCodigo)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpdo")
    a_Valores = Array(ps_ClsPlanilla, s_Periodo)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plperiodo", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fPeriodoPago.dcaRegistro, fPeriodoPago.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fPeriodoPago.dcaRegistro.Recordset.EOF And fPeriodoPago.dcaRegistro.Recordset.BOF) Or fPeriodoPago.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fPeriodoPago.dcaRegistro.Recordset.Find ("codpdo >= '" & s_Periodo & "'")
      If fPeriodoPago.dcaRegistro.Recordset.EOF Then fPeriodoPago.dcaRegistro.Recordset.MoveLast
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
   Case 0: fPeriodoPago.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fPeriodoPago.dcaRegistro.Recordset.BOF Then fPeriodoPago.dcaRegistro.Recordset.MovePrevious
           If fPeriodoPago.dcaRegistro.Recordset.BOF Then fPeriodoPago.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fPeriodoPago.dcaRegistro.Recordset.EOF Then fPeriodoPago.dcaRegistro.Recordset.MoveNext
           If fPeriodoPago.dcaRegistro.Recordset.EOF Then fPeriodoPago.dcaRegistro.Recordset.MoveLast
   Case 3: fPeriodoPago.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha final debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  If Not ((dtpFechas(2) >= dtpFechas(0)) And (dtpFechas(2) <= dtpFechas(1))) Then Beep: MsgBox "Fecha de pago debe ser mayor o igual a la fecha Inicial; menor o igual a la fecha Final", vbExclamation: dtpFechas(2).SetFocus: Exit Sub
  If txtFecha <> ps_Anyo Then Beep: MsgBox "Año debe ser del periodo activo", vbExclamation: txtFecha.SetFocus: Exit Sub
  If cmbPeriodo = "" Then Beep: MsgBox "Mes debe ser dentro del rango de las fechas", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
  If Not (Left(cmbPeriodo, 2) >= Mid$(dtpFechas(0), 4, 2) And Left(cmbPeriodo, 2) <= Mid$(dtpFechas(1), 4, 2)) Then Beep: MsgBox "Mes debe ser dentro del rango de las fechas", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
  If cmbtipo = "" Then Beep: MsgBox "Seleccione el Tipo de Periodo de Pago", vbExclamation: cmbtipo.SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, s_Estado_Ina, s_Estado_Act)
  
  ' Validación de que la fecha limite no sea menor a la del perido de pago que se desea Crear.
  ' Evita que un cambio en la fecha del sistema engañe al sistema.
  If Flag_RestringeSistema = "RESTRINGIR" Then
   If Valida_LicenciaUso(ps_Anyo, ps_Fecha_LimiteProc, Left(cmbPeriodo, 2), txtFecha.Text) = False Then
      MsgBox "Periódo de Pago no puede ser registrado" & Chr(13) & "Se requiere Actualización de Componentes" & Chr(13) & "Por favor comuniquese con el personal de Sistemas.", vbInformation, lblTitle.Caption
      Exit Sub
   End If
  End If
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Periodo = txtCodigo
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpdo", "despdo", "tpopdo", "fechaini", "fechafin", "fechapago", "anopdo", "mespdo", "estadopdo", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, txtCodigo, Trim$(txtDescripcion), Left$(cmbtipo, 1), Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), Format(dtpFechas(2), s_FmtFechMysql_0), Trim$(txtFecha), Trim(Left(cmbPeriodo, 2)), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpdo")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plperiodo", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plperiodo", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fPeriodoPago.dcaRegistro, fPeriodoPago.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fPeriodoPago.dcaRegistro.Recordset.Find ("codpdo='" & s_Periodo & "'")
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
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 5220: Me.Width = 7830
  Me.Left = 1080: Me.Top = 1500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Periodos de Pago"
  lblTitle = "Periodos de Pago"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fPeriodoPago.Tag

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
  l_ExistRecord = (fPeriodoPago.dcaRegistro.Recordset.EOF Or fPeriodoPago.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fPeriodoPago.dcaRegistro.Recordset!codpdo
  
  ' Configuro los listados, datos adicionales
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index

  ' Carga los datos en el formulario
  ShowScreen

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

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
Private Sub txtFecha_GotFocus()
  gdl_Procedure.MarcaGet txtFecha
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
