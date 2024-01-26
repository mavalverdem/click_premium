VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPPDTPcp 
   Caption         =   "[título]"
   ClientHeight    =   3390
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CmnDlgUbica 
      Left            =   225
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   893
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2573
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   767
      _Version        =   393216
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin ComctlLib.ProgressBar pgbEtapa1 
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   4
      Top             =   450
      Width           =   1635
   End
End
Attribute VB_Name = "frmPPDTPcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCOCpbPer As ADODB.Recordset
Public porstCOCpbPvs As ADODB.Recordset
Public porstCOCpbCan As ADODB.Recordset
Public pbNuevo As Boolean
Public pcNroCpb As String

Private Sub Form_Activate()
   LblProces.Visible = False
   cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo Err
   
   Dim dnContador As Integer
 
   cmdAceptar.Enabled = False
   cmdSalir.Enabled = False
   LblProces.Visible = True
   pgbEtapa1.Value = 0

  'Declaración de Variables.
   
  'Abrir Tablas.
   Set pocnnMain = New ADODB.Connection
   Set porstCOCpbPer = New ADODB.Recordset
   Set porstCOCpbPvs = New ADODB.Recordset
   Set porstCOCpbCan = New ADODB.Recordset

   With pocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG  & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstCOCpbPvs
      .ActiveConnection = pocnnMain
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
   End With
   With porstCOCpbCan
      .ActiveConnection = pocnnMain
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
   End With
   With porstCOCpbPer
      .ActiveConnection = pocnnMain
      .Source = "SELECT b.RucAux, b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, a.SerDoc, a.NroDoc, "
      .Source = .Source & "a.FeEDoc, a.ImpMN, b.TpoPer, CodDro, NroCpb "
      .Source = .Source & "FROM CoCpbDet a "
      .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux "
      .Source = .Source & "LEFT JOIN TGAuxNat c ON b.codemp=c.codemp AND b.CodAux=c.CodAux "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND a.CodTDc='" & gsCodTDc_Pcp & "'"
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
      .Open
   End With
'   pocnnMain.BeginTrans                'INICIA TRANSACCION.
 
  'Etapa1 : Generando Texto segun lectura de Tabla.
   
   dnContador = 0
   pgbEtapa1.Min = 0
''   pgbEtapa1.Max = 4
   pgbEtapa1.Value = pgbEtapa1.Min
   
   ppEtapa_01
   
   porstCOCpbPer.Close
   pocnnMain.Close
   Set porstCOCpbPvs = Nothing
   Set porstCOCpbCan = Nothing
   Set porstCOCpbPer = Nothing
   Set pocnnMain = Nothing
   
   MsgBox TEXT_8008, vbInformation
   cmdAceptar.Enabled = True
   cmdSalir.Enabled = True
   cmdSalir.SetFocus
   
   Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub ppEtapa_01()   ' Generacion de Texto en File
   Dim dnContador As Integer
   Dim dsTexto, dsFile As String
   dnContador = 0
   pgbEtapa1.Min = 0
   'Open "C:\Owl-paqu\Angel.TXT" For Output As #1
   dsFile = "0633" & gsRUCEmp & gsAnoAct & gsMesAct & ".TXT"
   CmnDlgUbica.FileName = dsFile
   CmnDlgUbica.ShowSave
   Open dsFile For Output As #1
   Do
      With porstCOCpbPer
         If .RecordCount = 0 Then
            Exit Do
         End If
         .MoveFirst
         pgbEtapa1.Max = .RecordCount
         pgbEtapa1.Value = pgbEtapa1.Min
         Do
            With porstCOCpbCan
               .Source = "SELECT CodAux, CodTDc, SerDoc, NroDoc, TpoGnr "
               .Source = .Source & "FROM CoCpbDet "
               .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
               .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
               .Source = .Source & "AND CodDro='" & porstCOCpbPer!CodDro & "' "
               .Source = .Source & "AND NroCpb='" & porstCOCpbPer!NroCpb & "' "
               .Source = .Source & "AND TpoPvs='" & TPOPVS_CAN & "' "
               .Source = .Source & "AND MesPvs='" & gsMesAct & "'"
               .Open
               If .RecordCount > 0 Then
                  .MoveFirst
                  Do
                     If !TpoGnr <> TPOGNR_DCA And !TpoGnr <> TPOGNR_DST Then
                        With porstCOCpbPvs
                           .Source = "SELECT CodTDc, SerDoc, NroDoc, FeEDoc, ImpMN, TpoGnr "
                           .Source = .Source & "FROM CoCpbDet "
                           .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
                           .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
                           .Source = .Source & "AND CodAux='" & porstCOCpbCan!CodAux & "' "
                           .Source = .Source & "AND CodTDc='" & porstCOCpbCan!CodTDc & "' "
                           .Source = .Source & "AND SerDoc='" & porstCOCpbCan!SerDoc & "' "
                           .Source = .Source & "AND NroDoc='" & porstCOCpbCan!NroDoc & "' "
                           .Source = .Source & "AND TpoPvs='" & TPOPVS_PVS & "'"
                           .Open
                           If .RecordCount > 0 Then
                              .MoveFirst
                              Do
                                 If !TpoGnr <> TPOGNR_DCA And !TpoGnr <> TPOGNR_DST Then
                                    dsTexto = Trim(porstCOCpbPer!RucAux) & "|"
                                    dsTexto = dsTexto & IIf(porstCOCpbPer!TpoPer = TPOPER_JUR, Trim(porstCOCpbPer!RazAux), "") & "|"
                                    dsTexto = dsTexto & IIf(porstCOCpbPer!TpoPer = TPOPER_JUR, "", Trim(porstCOCpbPer!ApePatAux)) & "|"
                                    dsTexto = dsTexto & IIf(porstCOCpbPer!TpoPer = TPOPER_JUR, "", Trim(porstCOCpbPer!ApeMatAux)) & "|"
                                    dsTexto = dsTexto & IIf(porstCOCpbPer!TpoPer = TPOPER_JUR, "", Trim(porstCOCpbPer!NomAux)) & "|"
                                    dsTexto = dsTexto & Trim(porstCOCpbPer!SerDoc) & "|"
                                    dsTexto = dsTexto & Right(Trim(porstCOCpbPer!NroDoc), 8) & "|"
                                    dsTexto = dsTexto & Format(porstCOCpbPer!FeEDoc, "dd/mm/yyyy") & "|"
                                    dsTexto = dsTexto & Trim(Str(porstCOCpbPer!ImpMN)) & "|"
                                    dsTexto = dsTexto & !CodTDc & "|"
                                    dsTexto = dsTexto & !SerDoc & "|"
                                    dsTexto = dsTexto & !NroDoc & "|"
                                    dsTexto = dsTexto & Format(!FeEDoc, "dd/mm/yyyy") & "|"
                                    dsTexto = dsTexto & !ImpMN & "|"
                                    Print #1, dsTexto
                                 End If
                                 .MoveNext
                              Loop Until .EOF
                           End If
                           .Close
                        End With
                     End If
                     .MoveNext
                  Loop Until .EOF
               End If
               .Close
            End With
''            Write #1, SUBSTR(dsTexto, 2, Len(dsTexto) - 2)
            dnContador = dnContador + 1
            pgbEtapa1.Value = dnContador
            .MoveNext
         Loop Until .EOF
      End With
      Exit Do
   Loop
   Close #1
End Sub

