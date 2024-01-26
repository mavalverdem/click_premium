VERSION 5.00
Begin VB.Form frmOCieMes 
   Caption         =   "[Entidad]"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1403
      ScaleHeight     =   690
      ScaleWidth      =   1575
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5280
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
         Picture         =   "frmOCieMes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
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
         Picture         =   "frmOCieMes.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   4935
      Left            =   60
      ScaleHeight     =   4875
      ScaleWidth      =   4215
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   4275
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         Caption         =   "Apertura"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   525
         Width           =   3255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   12
         Left            =   3360
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   4140
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   2680
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1140
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   2680
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   2680
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1740
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   4
         Left            =   2680
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   2680
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2340
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   2680
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   7
         Left            =   2680
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2940
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   8
         Left            =   2680
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   9
         Left            =   2680
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3540
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   10
         Left            =   2680
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   11
         Left            =   2680
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4140
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesHpr 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   2680
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   2000
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Enero"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   11
         Left            =   3360
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3540
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   9
         Left            =   3360
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2940
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2340
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1740
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1140
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   11
         Left            =   2000
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4140
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   10
         Left            =   2000
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   9
         Left            =   2000
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3540
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   8
         Left            =   2000
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   7
         Left            =   2000
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2940
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   2000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   2000
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2340
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   4
         Left            =   2000
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   2000
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1740
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   2000
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesVta 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   2000
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1140
         Width           =   255
      End
      Begin VB.CheckBox chkCieMesDro 
         Alignment       =   1  'Right Justify
         Caption         =   "Cierre"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4455
         Width           =   3255
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Diciembre"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4140
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Noviembre"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Octubre"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Setiembre"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Agosto"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Julio"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Junio"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Mayo"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Abril"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Marzo"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkCieMesCpr 
         Alignment       =   1  'Right Justify
         Caption         =   "Febrero"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   1200
         X2              =   3780
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lblTexto 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Diario"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   54
         Top             =   165
         Width           =   540
      End
      Begin VB.Label lblTexto 
         Caption         =   "Honorarios"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   2475
         TabIndex        =   55
         Top             =   165
         Width           =   975
      End
      Begin VB.Label lblTexto 
         Caption         =   "Ventas"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   1905
         TabIndex        =   56
         Top             =   165
         Width           =   615
      End
      Begin VB.Label lblTexto 
         Caption         =   "Compras"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   57
         Top             =   165
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmOCieMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstMain As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgOrde As String
Dim dnContador As Integer

Private Sub Form_Load()
   psConnStrgSele = "SELECT codemp, pdoano, MesCie, IndCpr, IndVta, IndHpr, IndCpb "
   psConnStrgSele = psConnStrgSele & "FROM COCieMes "
   psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
   psConnStrgSele = psConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
   psConnStrgOrde = "ORDER BY 1"
   Set pocnnMain = New ADODB.Connection
   Set porstMain = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
   End With
   With porstMain
      .ActiveConnection = pocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      '.Properties("Unique Table").Value = "Mescie"
   End With
   
   For dnContador = 0 To 11
      chkCieMesCpr(dnContador).Value = 0
      chkCieMesVta(dnContador).Value = 0
      chkCieMesHpr(dnContador).Value = 0
      chkCieMesDro(dnContador).Value = 0
   Next
   chkCieMesDro(12).Value = 0
   chkCieMesDro(13).Value = 0

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(4, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Compras", "Ventas", "Honorarios", "Diario")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Purchases", "Sales", "Feeds", "Journal")
  Next nElemento
  For nElemento = 0 To 11
    If gsIdioma = NvlUsr_Sup Then
      chkCieMesCpr(nElemento).Caption = Choose(nElemento + 1, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
    Else
      chkCieMesCpr(nElemento).Caption = Choose(nElemento + 1, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    End If
  Next nElemento
  chkCieMesDro(0).Caption = Choose(gsIdioma, "Apertura", "Opening")
  chkCieMesDro(13).Caption = Choose(gsIdioma, "Cierre", "Closing")
  CaptionBotones Me, True, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
   
   With porstMain
      .MoveFirst
      chkCieMesDro(0).Value = !IndCpb
      
      .MoveNext
      Do While Not .EOF
         If (!MesCie > 12) Then
            Exit Do
         End If
         chkCieMesCpr(!MesCie - 1).Value = !IndCpr
         chkCieMesVta(!MesCie - 1).Value = !IndVta
         chkCieMesHpr(!MesCie - 1).Value = !IndHpr
         chkCieMesDro(!MesCie).Value = !IndCpb
         .MoveNext
      Loop
      
      chkCieMesDro(!MesCie).Value = !IndCpb
   End With
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo Err
   
   Dim dnPos As Integer
    
   'Autor: Luis Almeyda
   pocnnMain.BeginTrans                'INICIA TRANSACCION.
   With porstMain
      dnPos = 0
      .MoveFirst
      Do While Not .EOF
          If (!MesCie = "00") Or (!MesCie = "13") Then
              !IndCpr = True
              !IndVta = True
              !IndHpr = True
              !IndCpb = chkCieMesDro(dnPos).Value
          Else
              !IndCpr = chkCieMesCpr(dnPos - 1).Value
              !IndVta = chkCieMesVta(dnPos - 1).Value
              !IndHpr = chkCieMesHpr(dnPos - 1).Value
              !IndCpb = chkCieMesDro(dnPos).Value
          End If
          .Update
          .MoveNext
          dnPos = dnPos + 1
      Loop
   End With
   pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
   
   gpCieMes
   
   Unload Me
   
   Exit Sub
Err:
   pocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub
