VERSION 5.00
Begin VB.Form frmOVar 
   Caption         =   "[Entidad]"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   300
      ScaleHeight     =   555
      ScaleWidth      =   3735
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Width           =   3795
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2880
         TabIndex        =   0
         Text            =   "12.12"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% Impuesto General a las Ventas:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   2460
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1440
      ScaleHeight     =   690
      ScaleWidth      =   1575
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1575
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   800
         Picture         =   "frmOVar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   60
         Picture         =   "frmOVar.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pocnnMain As ADODB.Connection
Private porstTGCfg As ADODB.Recordset

Private Sub Form_Load()
  'Abrir Tablas.
   Set pocnnMain = New ADODB.Connection
   Set porstTGCfg = New ADODB.Recordset

   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsRutBDS & gsNomBDS
      .Open
   End With
   With porstTGCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT PctIGV, " _
              & "UsrMdf_IGV, FyHMdf_IGV " _
              & "FROM TGCfg"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   
   txtDato(0).MaxLength = 5
   
   txtDato(0).Text = porstTGCfg!PctIGV
   txtDato(0).Tag = porstTGCfg!PctIGV
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err
   
   pocnnMain.BeginTrans                'INICIA TRANSACCION.
 
   With porstTGCfg
      If txtDato(0).Text <> txtDato(0).Tag Then
         !PctIGV = txtDato(0).Text
         !UsrMdf_IGV = gsAbvUsr
         !FyHMdf_IGV = Now
         .Update
      End If
   End With
   
   pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
         
   gnPctIGV = CDec(porstTGCfg!PctIGV)
   
'   MsgBox TEXT_8008, vbInformation
   cmdSalir.SetFocus
   
   Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

