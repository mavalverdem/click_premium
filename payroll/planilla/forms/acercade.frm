VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fAcercaDe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Click Premium"
   ClientHeight    =   4515
   ClientLeft      =   2370
   ClientTop       =   2025
   ClientWidth     =   6570
   Icon            =   "acercade.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4515
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame frmCuadro 
      Height          =   4500
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   6480
      _Version        =   65536
      _ExtentX        =   11430
      _ExtentY        =   7937
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
      ShadowStyle     =   1
      Begin VB.PictureBox pctAutoriza 
         Height          =   825
         Left            =   1470
         ScaleHeight     =   765
         ScaleWidth      =   4815
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2310
         Width           =   4875
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   135
            TabIndex        =   3
            Top             =   90
            Width           =   45
         End
         Begin VB.Label lblEmpresa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   135
            TabIndex        =   2
            Top             =   420
            Width           =   45
         End
      End
      Begin Threed.SSFrame frmLinea 
         Height          =   30
         Left            =   90
         TabIndex        =   4
         Top             =   3315
         Width           =   6240
         _Version        =   65536
         _ExtentX        =   11007
         _ExtentY        =   53
         _StockProps     =   14
         Caption         =   "SSFrame1"
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
      Begin Threed.SSCommand cmdAceptar 
         Height          =   345
         Left            =   4680
         TabIndex        =   5
         Top             =   3495
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "&Aceptar"
      End
      Begin Threed.SSCommand cmdActualizar 
         Height          =   345
         Left            =   3240
         TabIndex        =   11
         Top             =   1200
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Actualizar"
      End
      Begin VB.Image imgAcercaDe 
         Height          =   3105
         Left            =   105
         Stretch         =   -1  'True
         Top             =   165
         Width           =   1290
      End
      Begin VB.Label lblSoftware 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System's para Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1455
         TabIndex        =   10
         Top             =   165
         Width           =   2010
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 21.60       02/12/2021"
         Height          =   195
         Left            =   1455
         TabIndex        =   9
         Top             =   405
         Width           =   2160
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright© 2008-2022  AG Business Corporation"
         Height          =   195
         Left            =   1470
         TabIndex        =   8
         Top             =   630
         Width           =   3450
      End
      Begin VB.Label lblAutoriza 
         AutoSize        =   -1  'True
         Caption         =   "Se autoriza el uso de este Producto a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1455
         TabIndex        =   7
         Top             =   2010
         Width           =   3540
      End
      Begin VB.Label lblAdvertencia 
         BackStyle       =   0  'Transparent
         Caption         =   $"acercade.frx":000C
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   90
         TabIndex        =   6
         Top             =   3465
         Width           =   4485
      End
   End
End
Attribute VB_Name = "fAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                     ' Declarar variable antes de usarla
Private Sub cmdAceptar_Click()
  Unload Me
End Sub

Private Sub cmdActualizar_Click()
  Dim vActualizador As String
  
  vActualizador = "c:\Archivos de Programa\Personal y Planilla\Actualizador.exe"
  
  If FileOrDirExists(vActualizador) Then ' me aseguro que exista el archivo
    'ejecute el programa y todo lo demas
    Kill ("c:\Archivos de Programa\Personal y Planilla\Actualizador.exe")
    FileCopy "\\" & Quickref.sActualizador & "\usuarios\planilla\Actualizador.exe", "c:\Archivos de Programa\Personal y Planilla\Actualizador.exe"
  Else
    'copiar el actualizador
    FileCopy "\\" & Quickref.sActualizador & "\usuarios\planilla\Actualizador.exe", "c:\Archivos de Programa\Personal y Planilla\Actualizador.exe"
  End If
  
  If Val(ReadServerIni("VERSION", "VERSION")) > Val(Quickref.VERSION) Then
    Dim rp As String
    Dim mensaje As String
    mensaje = "Hay una version: " & Val(ReadServerIni("VERSION", "VERSION")) & " superior a: " & Quickref.VERSION & ", desea actualizar? " & "\\" & Quickref.sActualizador & "\usuarios\planilla\planilla.EXE"
    rp = MsgBox(mensaje, vbQuestion + vbYesNo, "Actualizar")
    If rp = vbYes Then
      Shell ps_PathSystem & "\Actualizador.exe", vbNormalFocus
      End
    Else
      'Si la Version es la misma continua proceso normal
    End If
  Else
    MsgBox "No Hay Actualizaciones por el Momento"
  End If
End Sub

Private Sub Form_Load()
  Dim s_Archivo As String, s_ToolText As String
  
  gdl_Procedure.CentraFormulario Me
  
  ' Cargo el icono de la ventana
  Me.Caption = "Acerca de " & ps_NomSistema
  Me.Icon = LoadPicture()
  s_Archivo = gdl_Procedure.ps_PathImagen & "acercade.ico"
  If dir$(s_Archivo, vbNormal) <> "" Then
      Me.Icon = LoadPicture(s_Archivo)
  End If
    
  ' Verifico que exista el Icono de Seguridad
  imgAcercaDe.Picture = LoadPicture()
  s_Archivo = gdl_Procedure.ps_PathImagen & "logo acercade.bmp"
  If dir$(s_Archivo, vbNormal) <> "" Then
      imgAcercaDe.Picture = LoadPicture(s_Archivo)
  End If
  imgAcercaDe.Refresh
  
  lblSoftware = ps_NomSistema & " para Windows"
  lblPersona = ps_Licencia
  lblEmpresa = ps_Licencia

End Sub
