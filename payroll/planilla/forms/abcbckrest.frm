VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form fAbcBckRest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7350
   Begin TabDlg.SSTab tabRegister 
      Height          =   5820
      Left            =   15
      TabIndex        =   0
      Top             =   585
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   10266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
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
      TabCaption(0)   =   "Backup/Restore"
      TabPicture(0)   =   "abcbckrest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmCuadro(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmCuadro(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "mensaje"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "framearchivos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame framearchivos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         TabIndex        =   2
         Top             =   660
         Width           =   6855
         Begin VB.DriveListBox driveexe 
            BackColor       =   &H80000013&
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   3195
         End
         Begin VB.DirListBox direxe 
            BackColor       =   &H80000013&
            Height          =   990
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   3195
         End
         Begin VB.TextBox ruta 
            BackColor       =   &H80000013&
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   6615
         End
         Begin VB.FileListBox exes 
            BackColor       =   &H80000013&
            Height          =   1065
            Left            =   3360
            Pattern         =   "mysqldump.exe;mysql.exe"
            TabIndex        =   3
            Top             =   840
            Width           =   3375
         End
      End
      Begin VB.TextBox mensaje 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   4680
         Width           =   6855
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1725
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   3043
         _StockProps     =   14
         Caption         =   " Ubicación   Nota: Utilizar Directorios sin espacios en blanco"
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
         Begin VB.DriveListBox drive 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3195
         End
         Begin VB.DirListBox dir 
            Height          =   990
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   3195
         End
         Begin VB.FileListBox files 
            BackColor       =   &H00E0E0E0&
            Height          =   1260
            Left            =   3360
            Pattern         =   "*.bsql"
            TabIndex        =   8
            Top             =   240
            Width           =   3375
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   50
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   1085
         _StockProps     =   14
         Caption         =   " Opción "
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
         Begin Threed.SSOption opt 
            Height          =   200
            Index           =   0
            Left            =   230
            TabIndex        =   12
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Backup"
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
            Value           =   -1  'True
         End
         Begin Threed.SSOption opt 
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Restore"
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
         End
         Begin VB.Label nombre 
            Caption         =   "RUCddmmyyyy.bsql"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   14
            Top             =   120
            Width           =   2895
         End
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7350
      _Version        =   65536
      _ExtentX        =   12965
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
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   6840
         TabIndex        =   16
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
         Picture         =   "abcbckrest.frx":001C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6360
         TabIndex        =   17
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcbckrest.frx":0038
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BACKUP - RESTAURAR BASE DE DATOS ( Servidor)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   105
         TabIndex        =   18
         Top             =   105
         Width           =   6165
      End
   End
End
Attribute VB_Name = "fAbcBckRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAction_Click(Index As Integer)
  Dim arch As String
  
  'Dim oExecuteCmd As Object
  ' Instancio objetode comando de ejecuccion
  'Set oExecuteCmd = CreateObject("WSCript.shell")
  'oExecuteCmd.Run
  'Set oExecuteCmd = Nothing
  
  Select Case Index
   Case 0
    If frmCuadro(1).Visible = False Then Exit Sub
    If opt(0).Value = True Then
      If Len(dir.path) = 3 Then
        arch = drive.drive
      Else
        arch = dir.path
      End If
      arch = arch & "\" & ps_RucEmpresa & Year(Date) & Month(Date) & Day(Date) & ".bsql"
      TipodeProgreso = 1
      IntervalodeTiempo = 300
      labelprogreso = "Generando                                                          Copia de Seguridad"
      mensaje.Text = "cmd /K cd\ & cd " & ruta.Text & " & mysqldump --host=" & ps_Servidor & " --user=" & ps_UserId & " --password=" & ps_Password & " " & ps_DataBase & " > " & arch & " & exit "
      Shell "cmd /K cd\ & cd " & ruta.Text & " & mysqldump --host=" & ps_Servidor & " --user=" & ps_UserId & " --password=" & ps_Password & " " & ps_DataBase & " > " & arch & " & exit "

      Progreso.Show vbModal
    Else
      If files.FileName = "" Then
        MsgBox ("Tiene que Seleccionar un Archivo")
        Exit Sub
      End If
    
      If Left(files.FileName, 11) <> ps_RucEmpresa Then
        MsgBox ("Verificar Archivo, Tiene que tener el mismo Ruc de la Empresa al que quiere Restaurar")
        Exit Sub
      End If
      TipodeProgreso = 1
      IntervalodeTiempo = 400
      labelprogreso = "Restaurando                                                         Copia de Seguridad"
      mensaje.Text = "cmd /K cd\ & cd " & ruta.Text & " & mysql --host=" & ps_Servidor & " --user=" & ps_UserId & " --password=" & ps_Password & " " & ps_DataBase & " < " & dir.path & IIf(Len(dir.path) = 3, "", "\") & files.FileName & " & exit "
      Shell "cmd /K cd\ & cd " & ruta.Text & " & mysql --host=" & ps_Servidor & " --user=" & ps_UserId & " --password=" & ps_Password & " " & ps_DataBase & " < " & dir.path & IIf(Len(dir.path) = 3, "", "\") & files.FileName & " & exit "
      Progreso.Show vbModal
    End If
   Case 1
    Unload Me: Exit Sub
  End Select

End Sub

Private Sub exes_PathChange()

  If exes.ListCount = 2 Then
    frmCuadro(1).Visible = True
  Else
    frmCuadro(1).Visible = False
  End If

End Sub

Private Sub Form_Load()
  'Establece posición y titulo del formulario
  Me.Height = 6900: Me.Width = 7440
  Me.Left = 1080: Me.Top = 200
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(2, 1) = "proceso": aElemento(2, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 1
    aElemento(n_Index, 1) = Choose(n_Index + 1, "link", "cancelar")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Procesar ", "Cancelar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(1).Cancel = True
  
  On Error GoTo Error
  dir.path = "C:\"
  ruta.Text = ps_MysqlExe
  driveexe.drive = "C:\"
  
  On Error GoTo Error
  direxe.path = ps_MysqlExe
  If exes.ListCount = 2 Then
    frmCuadro(1).Visible = True
  Else
    frmCuadro(1).Visible = False
  End If
  files.Visible = False
Error:

End Sub
Private Sub Drive_Change()
  On Error GoTo Err_
  dir.path = drive.drive
  Exit Sub
Err_:   drive.drive = "C:\"
End Sub
Private Sub Dir_Change()
  On Error GoTo Err_
  files.path = dir.path
  Exit Sub
Err_:   drive.drive = "C:\"
End Sub
Private Sub Driveexe_Change()
  On Error GoTo Err_
  direxe.path = driveexe.drive
  Exit Sub
Err_:   driveexe.drive = "C:\"
End Sub
Private Sub Direxe_Change()
  On Error GoTo Err_
  exes.path = direxe.path
  ruta.Text = direxe.path
  Exit Sub
Err_:   driveexe.drive = "C:\"
End Sub
Private Sub opt_Click(Index As Integer, Value As Integer)
  Select Case Index
   Case 0
    files.Visible = False
    mensaje.Text = ""
   Case 1
    files.Visible = True
    mensaje.Text = ""
  End Select
End Sub
