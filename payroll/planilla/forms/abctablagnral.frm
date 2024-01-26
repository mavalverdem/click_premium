VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fAbcTablaGeneral 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7200
   Icon            =   "abctablagnral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   7200
   Begin TabDlg.SSTab tabRegister 
      Height          =   4275
      Left            =   75
      TabIndex        =   47
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   7541
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
      TabPicture(0)   =   "abctablagnral.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDescripcion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbTipo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDefault"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtDefault 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1340
         TabIndex        =   5
         Top             =   990
         Width           =   1425
      End
      Begin VB.ComboBox cmbTipo 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "abctablagnral.frx":0028
         Left            =   4230
         List            =   "abctablagnral.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   990
         Width           =   1800
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         TabIndex        =   3
         Top             =   630
         Width           =   4680
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1340
         TabIndex        =   1
         Top             =   270
         Width           =   980
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2400
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   1455
         Width           =   5865
         _Version        =   65536
         _ExtentX        =   10345
         _ExtentY        =   4233
         _StockProps     =   14
         Caption         =   " Valores por Periodos "
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
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   4215
            TabIndex        =   32
            Top             =   1995
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   4215
            TabIndex        =   30
            Top             =   1665
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   4215
            TabIndex        =   28
            Top             =   1335
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   4215
            TabIndex        =   26
            Top             =   1005
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   4215
            TabIndex        =   24
            Top             =   675
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   4215
            TabIndex        =   22
            Top             =   345
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   1250
            TabIndex        =   20
            Top             =   1995
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   1250
            TabIndex        =   18
            Top             =   1665
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   1250
            TabIndex        =   16
            Top             =   1335
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   1250
            TabIndex        =   14
            Top             =   1005
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   1250
            TabIndex        =   12
            Top             =   675
            Width           =   1425
         End
         Begin VB.TextBox txtValor 
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1250
            TabIndex        =   10
            Top             =   345
            Width           =   1425
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Diciembre :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   3120
            TabIndex        =   31
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Noviembre :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   3120
            TabIndex        =   29
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Octubre :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   3120
            TabIndex        =   27
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Setiembre :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   3120
            TabIndex        =   25
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Agosto :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   3120
            TabIndex        =   23
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Julio :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   3120
            TabIndex        =   21
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Junio :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   19
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Mayo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   17
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Abril :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   15
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Marzo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   13
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Febrero :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   11
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Enero :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   9
            Top             =   390
            Width           =   1005
         End
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Default :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3105
         TabIndex        =   6
         Top             =   1035
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
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo :"
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
      TabIndex        =   33
      Top             =   0
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Left            =   6450
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
         Picture         =   "abctablagnral.frx":002C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   35
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
         Picture         =   "abctablagnral.frx":0048
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
         TabIndex        =   36
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   37
      Top             =   4935
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Picture         =   "abctablagnral.frx":0064
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   39
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
         Picture         =   "abctablagnral.frx":0080
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   40
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
         Picture         =   "abctablagnral.frx":009C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   41
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
         Picture         =   "abctablagnral.frx":00B8
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4275
      Index           =   0
      Left            =   6435
      TabIndex        =   42
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   7541
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         Picture         =   "abctablagnral.frx":00D4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   45
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
         Picture         =   "abctablagnral.frx":00F0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   46
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
         Picture         =   "abctablagnral.frx":010C
      End
   End
End
Attribute VB_Name = "fAbcTablaGeneral"
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
'[
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

  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fTablasGeneral.dcaRegistro.Recordset!codtbl.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fTablasGeneral.dcaRegistro.Recordset!destbl.DefinedSize
    gdl_Procedure.EditText "AT", txtDefault, FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditCombo "PK", cmbtipo, -1, Me.Tag, False
    For n_Index = 0 To 11
      gdl_Procedure.EditText "AT", txtValor(n_Index), FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    Next n_Index
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fTablasGeneral.dcaRegistro.Recordset!codtbl, Me.Tag, True, fTablasGeneral.dcaRegistro.Recordset!codtbl.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fTablasGeneral.dcaRegistro.Recordset!destbl), Me.Tag, False, fTablasGeneral.dcaRegistro.Recordset!destbl.DefinedSize
    gdl_Procedure.EditText "AT", txtDefault, FormatNumber(fTablasGeneral.dcaRegistro.Recordset!valordefa, 2), Me.Tag, False, 18, vbRightJustify
    n_Index = IIf(fTablasGeneral.dcaRegistro.Recordset!tpotbl = "M", 0, 1)
    gdl_Procedure.EditCombo "AT", cmbtipo, n_Index, Me.Tag, False
    For n_Index = 0 To 11
      gdl_Procedure.EditText "AT", txtValor(n_Index), FormatNumber(fTablasGeneral.dcaRegistro.Recordset("valor" & Format((n_Index + 1), "00")), 2), Me.Tag, False, 18, vbRightJustify
    Next n_Index
  End If

End Sub
']
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   txtDescripcion.SetFocus
  End If
  If Index <> 1 Then Exit Sub
    
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & txtDescripcion & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim$(txtCodigo)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "pdoano", "codtbl")
    a_Valores = Array(ps_ClsPlanilla, ps_Anyo, txtCodigo)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("pltablabase", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fTablasGeneral.dcaRegistro, fTablasGeneral.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fTablasGeneral.dcaRegistro.Recordset.EOF And fTablasGeneral.dcaRegistro.Recordset.BOF) Or fTablasGeneral.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fTablasGeneral.dcaRegistro.Recordset.Find ("codtbl >= '" & s_Registro & "'")
      If fTablasGeneral.dcaRegistro.Recordset.EOF Then fTablasGeneral.dcaRegistro.Recordset.MoveLast
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

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fTablasGeneral.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fTablasGeneral.dcaRegistro.Recordset.BOF Then fTablasGeneral.dcaRegistro.Recordset.MovePrevious
           If fTablasGeneral.dcaRegistro.Recordset.BOF Then fTablasGeneral.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fTablasGeneral.dcaRegistro.Recordset.EOF Then fTablasGeneral.dcaRegistro.Recordset.MoveNext
           If fTablasGeneral.dcaRegistro.Recordset.EOF Then fTablasGeneral.dcaRegistro.Recordset.MoveLast
   Case 3: fTablasGeneral.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If cmbtipo = "" Then Beep: MsgBox "Seleccione Tipo de Valor " & lblTitle, vbExclamation: cmbtipo.SetFocus: Exit Sub
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = txtCodigo
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "pdoano", "codtbl", "destbl", "tpotbl", "valordefa", "valor01", "valor02", "valor03", "valor04", "valor05", "valor06", "valor07", "valor08", "valor09", "valor10", "valor11", "valor12", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, ps_Anyo, txtCodigo, gdl_Funcion.SacaEntRetApos(txtDescripcion), Left(cmbtipo, 1), CDec(txtDefault), CDec(txtValor(0)), CDec(txtValor(1)), CDec(txtValor(2)), CDec(txtValor(3)), CDec(txtValor(4)), CDec(txtValor(5)), CDec(txtValor(6)), CDec(txtValor(7)), CDec(txtValor(8)), CDec(txtValor(9)), CDec(txtValor(10)), CDec(txtValor(11)), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "pdoano", "codtbl")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("pltablabase", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("pltablabase", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fTablasGeneral.dcaRegistro, fTablasGeneral.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fTablasGeneral.dcaRegistro.Recordset.Find ("codtbl='" & s_Registro & "'")
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

  'Establece posición y titulo del formulario
  Me.Height = 5920: Me.Width = 7290
  Me.Left = 1080: Me.Top = 950
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Tablas Generales"
  lblTitle = "Tabla General"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operación del registro
  Me.Tag = fTablasGeneral.Tag
  
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
  l_ExistRecord = (fTablasGeneral.dcaRegistro.Recordset.EOF Or fTablasGeneral.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fTablasGeneral.dcaRegistro.Recordset!codtbl
  
  ' Adiciono los tipos de valores
  For n_Index = 0 To 1
    cmbtipo.AddItem Choose(n_Index + 1, "Monto", "Tasa")
  Next n_Index
  
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
Private Sub txtDefault_GotFocus()
  gdl_Procedure.MarcaGet txtDefault
End Sub
Private Sub txtDefault_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtDefault_Validate(Cancel As Boolean)
  txtDefault.Text = IIf(Not IsNumeric(txtDefault.Text), 0, txtDefault.Text)
  txtDefault.Text = FormatNumber(CDec(txtDefault.Text), 2)
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
Private Sub txtValor_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtValor(Index)
End Sub
Private Sub txtValor_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtValor_Validate(Index As Integer, Cancel As Boolean)
  txtValor(Index).Text = IIf(Not IsNumeric(txtValor(Index).Text), 0, txtValor(Index).Text)
  txtValor(Index).Text = FormatNumber(CDec(txtValor(Index).Text), 2)
End Sub
