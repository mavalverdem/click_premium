VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fConsultaVarios 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   4260
   Icon            =   "convarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   4260
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4260
      _Version        =   65536
      _ExtentX        =   7514
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
         Index           =   0
         Left            =   3705
         TabIndex        =   3
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
         Picture         =   "convarios.frx":000C
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
         TabIndex        =   4
         Top             =   120
         Width           =   3180
      End
   End
   Begin Threed.SSFrame frmCuadro 
      Height          =   3570
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   6297
      _StockProps     =   14
      Caption         =   " Parametro de Consulta "
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
      Begin MSComctlLib.TreeView tvwConsulta 
         Height          =   3105
         Left            =   240
         TabIndex        =   5
         Top             =   345
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5477
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   3270
         Index           =   1
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   270
         Width           =   4020
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   4215
      Width           =   4260
      _Version        =   65536
      _ExtentX        =   7514
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
   Begin MSComctlLib.ImageList imlGrafico 
      Left            =   3855
      Top             =   3630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image imgGrafico 
      Height          =   240
      Left            =   3975
      Top             =   3330
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "fConsultaVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
'[
Private Sub cmdCancel_Click(Index As Integer)
  Unload Me
End Sub
Private Sub Form_Load()
  Dim s_Archivo As String
  Dim o_Nodox As Node
  
  'Establece posición y titulo del formulario
  Me.Height = 5200: Me.Width = 4350
  Me.Left = 500: Me.Top = 500
  
  ' Titulo del formulario y panel
  s_TitleWindow = Me.Caption
  lblTitle = "Registro Varios"
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "proceso": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "cancelar"
  aElemento(0, 2) = "Cancelar Proceso de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdCancel, aElemento
  cmdCancel(0).Cancel = True
  '[
  ' Configuro los graficos de objeto
  For n_Index = 1 To 10
    s_Archivo = gdl_Procedure.ps_PathImagen & Choose(n_Index, "consulta", "padron", "inftraba", "vacacion", "hisfases", "explabor", "estudios", "datofami", "contrato", "fichapsn") & ".bmp"
    If dir$(s_Archivo, vbNormal) <> "" Then
      imgGrafico.Picture = LoadPicture(s_Archivo)
    End If
    imgGrafico.Refresh
    imlGrafico.ListImages.Add , , imgGrafico
  Next n_Index
  ' Agrego el objeto de graficos
  tvwConsulta.ImageList = imlGrafico
  tvwConsulta.Style = tvwTreelinesPictureText
  tvwConsulta.Indentation = 300
  
  ' Configuro el objeto de parametro
  Set o_Nodox = tvwConsulta.Nodes.Add(, , "prm", "Registros Varios", 1)
  
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op1", "Padrón de Empleados", 2)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op2", "Datos de Trabajos", 3)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op3", "Rol de Vacaciones", 4)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op4", "Remuneraciones", 5)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op5", "Experiencia Laboral", 6)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op6", "Estudios Realizados", 7)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op7", "Datos Familiares", 8)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op8", "Contratos", 9)
  Set o_Nodox = tvwConsulta.Nodes.Add("prm", tvwChild, "op9", "Ficha de Datos", 10)
  
  o_Nodox.EnsureVisible
  
  ']
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub

Private Sub tvwConsulta_NodeClick(ByVal Node As MSComctlLib.Node)
  Me.Tag = Node.Key
  If Me.Tag = "prm" Then Exit Sub
  fSelPersonalCst.Show
End Sub
