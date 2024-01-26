VERSION 5.00
Begin VB.Form frmPeli 
   Caption         =   "[titulo]"
   ClientHeight    =   2895
   ClientLeft      =   4050
   ClientTop       =   1905
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   4665
   Begin VB.Frame fraProceso 
      Caption         =   "Eliminacion de Procesos Automaticos"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4215
      Begin VB.CheckBox chkProceso 
         Caption         =   "Diferencia de C&ambio"
         Height          =   200
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3880
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "&Asientos de Destino"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3880
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Asiento de Ap&ertura"
         Height          =   200
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3880
      End
      Begin VB.CheckBox chkProceso 
         Caption         =   "Asiento de C&ierrre"
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3880
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2460
      TabIndex        =   1
      Top             =   2085
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   780
      TabIndex        =   0
      Top             =   2085
      Width           =   1215
   End
End
Attribute VB_Name = "frmPeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection

Private Sub cmdAceptar_Click()
  ' Validacionde mes cerrado
  If gbCieCpb Then MsgBox TEXT_9016, vbCritical: cmdSalir.SetFocus: Exit Sub
  
  pocnnMain.BeginTrans                'INICIA TRANSACCION.
  'Paso 1 : Elimino los comprobantes de ajuste del mes
  If chkProceso(1).Value = vbChecked Then
    pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DST) & " And MesPvs='" & gsMesAct & "'"
    If gnProDestino = NvlUsr_Sup Then
      pocnnMain.Execute "DELETE FROM cocpbdet WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr='" & TPOGNR_DST & "' And MesPvs='" & gsMesAct & "'"
    End If
  End If
  'Paso 2: Elimino los asientos de destino
  If chkProceso(2).Value = vbChecked Then
    pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DCA) & " AND MesPvs='" & gsMesAct & "'"
  End If
  'Paso 3: Elimino los asientos de Apertura
  If chkProceso(3).Value = vbChecked Then
    pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_APE) & " AND MesPvs='" & gsMesAct & "'"
  End If
  'Paso 4: Elimino los asientos de Cierre
  If chkProceso(4).Value = vbChecked Then
    pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_CIE) & " AND MesPvs='" & gsMesAct & "'"
  End If
  
'ini 2015-08-21 Si Mayorizo o no . Estado Mayorizacion
  If chkProceso(2).Value = vbChecked Or chkProceso(3).Value = vbChecked Then
    fEstMayUpd
  End If
'fin 2015-08-21 Si Mayorizo o no . Estado Mayorizacion
  pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
  cmdSalir.SetFocus
'ini 2015-06-08 Si Mayorizo o no . Estado Mayorizacion
  MsgBox (TEXT_8008)
'fin 2015-06-08 Si Mayorizo o no . Estado Mayorizacion
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub txtDato_Change(Index As Integer)

End Sub

Private Sub Form_Load()
'Abrir Tablas.
   
   Set pocnnMain = New ADODB.Connection
   
   With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
   End With
   
   '[ Cargo los mensajes de botones
   ReDim aLabel(0, 0)
   fraProceso.Caption = Choose(gsIdioma, "Eliminación de Asientos Automaticos", "Eliminating of Automatic Entries")
   chkProceso(1).Caption = Choose(gsIdioma, "Asientos de &Destinos", "&Destination Entries")
   chkProceso(2).Caption = Choose(gsIdioma, "Diferencia de &Cambio", "Defference of &Exchange")
   chkProceso(3).Caption = Choose(gsIdioma, "Asientos de &Apertura", "&Opening Entries")
   chkProceso(4).Caption = Choose(gsIdioma, "Asientos de Ci&erre", "C&losing Entries")
   cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
   CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
  pocnnMain.Close
  Set pocnnMain = Nothing
End Sub
