VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmZCtrObliSunat 
   Caption         =   "Datos de la presentación"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
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
      Left            =   2160
      Picture         =   "frmZCtrObliSunat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   720
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   1200
      Picture         =   "frmZCtrObliSunat.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   720
   End
   Begin VB.TextBox txtDato 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   210
      Width           =   2310
   End
   Begin VB.TextBox txtDato 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   600
      Width           =   2310
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65994753
      CurrentDate     =   37102
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Declaración :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   945
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Constancia :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   630
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Presentación :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1005
      Width           =   1170
   End
End
Attribute VB_Name = "frmZCtrObliSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
'ini 2015-08-27/09-02 ctr obligac sunat
    Dim xx_ano_act As String
    Dim xx_mes_act As String
    Dim xx_fe_ante As Date
    xx_fe_ante = gfMesAnte(CDate("01/" + Format(Month(Now), "00") + "/" + Format(Year(Now), "0000")))
    'gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
'   xx_ano_act = Format(xx_fe_ante, "0000")
'    xx_mes_act = Format(xx_fe_ante, "00")
    xx_ano_act = Format(Year(xx_fe_ante), "0000")
    xx_mes_act = Format(Month(xx_fe_ante), "00")
'fin 2015-08-27/09-02 ctr obligac sunat

   Dim docnnMain As ADODB.Connection
   Set docnnMain = New ADODB.Connection
   
    Dim x_Sentencia As String
    x_Sentencia = "INSERT INTO " & gsNomBDC & fPunto & "tgCtrOblidet ("
    x_Sentencia = x_Sentencia & "codemp,pdotribu,"
    x_Sentencia = x_Sentencia & "CodDeclar, nroConsta,  fpresenta,"
    x_Sentencia = x_Sentencia & " usrcre, fyhcre"
    x_Sentencia = x_Sentencia & ")"
    x_Sentencia = x_Sentencia & " VALUES ("
    x_Sentencia = x_Sentencia & "'" + gsCodEmp & "'"
    'x_Sentencia = x_Sentencia & ",'" + gsAnoAct & gsMesAct & "'"
    x_Sentencia = x_Sentencia & ",'" + xx_ano_act & xx_mes_act & "'"
    x_Sentencia = x_Sentencia & ",'" + txtDato(0).Text & "'"
    x_Sentencia = x_Sentencia & ",'" + txtDato(1).Text & "'"
    x_Sentencia = x_Sentencia & "," + fDateFmt(dtpDato(0).Value) & " "
    x_Sentencia = x_Sentencia & ",'" + gsAbvUsr & "'"
    x_Sentencia = x_Sentencia & "," + fDateNow() & " "
    x_Sentencia = x_Sentencia & ")"
'dtpDato(3).Value
'    x_Sentencia = x_Sentencia & "CodEmp ='" & gsCodEmp & "' "
'    x_Sentencia = x_Sentencia & "AND PdoAno ='" & gsAnoAct & "' "
'    x_Sentencia = x_Sentencia & "AND MesCie='" & gsMesAct & "' "
   Set docnnMain = fCnnOpen(docnnMain)
   docnnMain.Execute x_Sentencia
'   With docnnMain
'    .CursorLocation = adUseClient
'    .ConnectionString = CONNSTRG & gsNomBDS
'    .Open
'    .Execute x_Sentencia
'   End With
   
fCnnClose docnnMain
   Unload Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    txtDato(0).MaxLength = 4
    txtDato(1).MaxLength = 25

  dtpDato(0).Value = Date

  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, _
    "Código PDT :", "Nro.Constancia :", "F.Presentación :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, _
    "code PDT:", "Num.Constancy :", "D.Submission:")
  Next nElemento
  
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, True, False, True, aLabel

End Sub
