VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPMayEst 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6240
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
         Left            =   240
         Picture         =   "frmPMayEst.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPMayEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public uocnnMain As ADODB.Connection
Public uorstMain As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgOrde As String

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   ppDatosGrid
End Sub

Private Sub Form_Load()
    Me.Caption = Choose(gsIdioma, "Estado de la Mayorizacion", "State Majorization") '"Estado de la Mayorizacion"
    Set uocnnMain = New ADODB.Connection
    Set uorstMain = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
'   psConnStrgSele = "SELECT CodDro, "
'   psConnStrgSele = psConnStrgSele & Choose(gsIdioma, "DetDro, ", "DetDrox, ")
'   psConnStrgSele = psConnStrgSele & Choose(gsIdioma, "DetDrox, ", "DetDro, ")
'   psConnStrgSele = psConnStrgSele & "codemp, pdoano, UsrCre, FyHCre, UsrMdf, FyHMdf,codlib "
'   psConnStrgSele = psConnStrgSele & "FROM CoDro "
'   psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
'   psConnStrgSele = psConnStrgSele & "AND pdoano='" & gsAnoAct & "'"
'   psConnStrgOrde = "ORDER BY 1"
   
   psConnStrgSele = "SELECT PdoAno, MesCie, indProcMay "
   psConnStrgSele = psConnStrgSele & "FROM CoCieMes "
   psConnStrgSele = psConnStrgSele & "WHERE codemp='" & gsCodEmp & "' "
   '2015-09-04 error bitacora y ctr mayoriza psConnStrgSele = psConnStrgSele & "AND pdoano='" & gsAnoAct & "' "
   psConnStrgSele = psConnStrgSele & "AND indProcMay <> 0 "
   psConnStrgOrde = "ORDER BY 2"
   
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      '.Properties("Unique Table").Value = "CODRO"
'      .Properties("Unique Table").Value = "a"
   End With
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   gpTUg_Resize Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
   uorstMain.Close
   uocnnMain.Close
   Set uorstMain = Nothing
   Set uocnnMain = Nothing

End Sub

Public Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
   'Year. Month. state
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Año", "Year")
'            .Item(dnNum).Width = 600
             .Item(dnNum).Width = 100 * (uorstMain.Fields("PdoAno").DefinedSize + 10)
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Mes", "Month")
'            .Item(dnNum).Width = 1500
             .Item(dnNum).Width = 100 * (uorstMain.Fields("MesCie").DefinedSize + 12)
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Estado", "State")
'            .Item(dnNum).Width = 1500
             .Item(dnNum).Width = 100 * (uorstMain.Fields("indProcMay").DefinedSize + 12)
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub


