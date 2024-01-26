VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPPDTDAOT 
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
      Left            =   45
      Top             =   2520
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
   Begin ComctlLib.ProgressBar PgBEtapa2 
      Height          =   345
      Left            =   225
      TabIndex        =   5
      Top             =   1575
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando Ventas"
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
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Top             =   1305
      Width           =   2355
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando Compras"
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
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   450
      Width           =   2355
   End
End
Attribute VB_Name = "frmPPDTDAOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public pocnnConf As ADODB.Connection
Public porstCOCprDoc As ADODB.Recordset
Public porstCOVtaDoc As ADODB.Recordset
Public porstTGEMP As ADODB.Recordset
Public pbNuevo As Boolean
Public pcNroCpb As String

Private Sub Form_Activate()
   LblProces(0).Visible = False
   LblProces(1).Visible = False
   cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo Err
   
   Dim dnContador As Integer
 
   cmdAceptar.Enabled = False
   cmdSalir.Enabled = False
   LblProces(0).Visible = True
   LblProces(1).Visible = False
   pgbEtapa1.Value = 0
   PgBEtapa2.Value = 0

  'Declaración de Variables.
   
  'Abrir Tablas.
   Set pocnnMain = New ADODB.Connection
   Set pocnnConf = New ADODB.Connection
   Set porstTGEMP = New ADODB.Recordset
   Set porstCOCprDoc = New ADODB.Recordset
   Set porstCOVtaDoc = New ADODB.Recordset

   With pocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG  & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With pocnnConf
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With porstTGEMP
      .ActiveConnection = pocnnConf
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
   End With
   With porstCOCprDoc
      .ActiveConnection = pocnnMain
      .Source = "SELECT b.Tpoper, b.RucAux,"
      .Source = .Source & " SUM(IF(d.SgnTDc=0,(a.ImpOGr_MN+a.ImpOGn_MN+a.ImpONG_MN+a.ImpExo_mn) * -1,"
      .Source = .Source & " (a.ImpOGr_MN+a.ImpOGn_MN+a.ImpONG_MN+a.ImpExo_mn))) AS Total,"
      .Source = .Source & " b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, a.CodAux"
      .Source = .Source & " FROM CoCprDoc a Left Join TGAux b On a.CodAux=b.CodAux"
      .Source = .Source & " LEFT Join TGAuxNat c On b.CodAux=c.CodAux"
      .Source = .Source & " LEFT JOIN TgTDc d ON a.CodTDc=d.CodTDc"
      .Source = .Source & " GROUP BY a.CodAux"
      .Source = .Source & " HAVING Total>" & CDec(gnImpUIT) * 3
      .Source = .Source & " ORDER BY a.CodAux"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
      .Open
      .Properties("Unique Table").Value = "COCprDoc"
   End With
   
'   pocnnMain.BeginTrans                'INICIA TRANSACCION.
 
  'Etapa1 : Generando Texto segun lectura de Tabla.
   
   dnContador = 0
   pgbEtapa1.Min = 0
''   pgbEtapa1.Max = 4
   pgbEtapa1.Value = pgbEtapa1.Min
   
   ppEtapa_01
   
   With porstCOVtaDoc
      .ActiveConnection = pocnnMain
      .Source = "SELECT b.Tpoper, b.RucAux," _
              & "SUM(IF(d.SgnTDc=0,(a.ImpOGr_MN+a.ImpExp_MN+a.ImpExo_mn) * -1," _
              & "(a.ImpOGr_MN+a.ImpExp_MN+a.ImpExo_mn))) AS Total," _
              & "b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, a.CodAux " _
              & "FROM CoVtaDoc a Left Join TGAux b On a.CodAux=b.CodAux" _
              & " Left Join TGAuxNat c On b.CodAux=c.CodAux" _
              & " LEFT JOIN TgTDc d ON a.CodTDc=d.CodTDc" _
              & " Group by a.CodAux " _
              & " HAVING Total>" & CDec(gnImpUIT) * 3
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
      .Open
      .Properties("Unique Table").Value = "COVtaDoc"
   End With
   LblProces(1).Visible = True
   
   ppEtapa_02
   
   porstCOCprDoc.Close
   porstCOVtaDoc.Close
   pocnnConf.Close
   pocnnMain.Close
   Set porstTGEMP = Nothing
   Set porstCOCprDoc = Nothing
   Set porstCOVtaDoc = Nothing
   Set pocnnConf = Nothing
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

Private Sub ppEtapa_01()   ' Generacion de Texto en File Costos
   Dim dnContador As Integer, n_Posicion As Integer
   Dim dsTexto, dsFile As String
   dnContador = 0
   pgbEtapa1.Min = 0
    With porstTGEMP
      .Source = "Select RucEmp From TGEMP Where CodEmp='" & gsCodEmp & "'"
      .Open
   End With
   dsFile = "Costos.TXT"
   CmnDlgUbica.FileName = dsFile
   CmnDlgUbica.ShowSave
   Open dsFile For Output As #1
   Do
      With porstCOCprDoc
         If .RecordCount = 0 Then
            Exit Do
         End If
         .MoveFirst
         pgbEtapa1.Max = .RecordCount
         pgbEtapa1.Value = pgbEtapa1.Min
         Do
            dnContador = dnContador + 1
            dsTexto = Trim(Str(dnContador)) & "|"
            dsTexto = dsTexto & "6|" & porstTGEMP!RUCEmp & "|"
            dsTexto = dsTexto & gsAnoAct & "|"
            dsTexto = dsTexto & IIf(!TpoPer = TPOPER_JUR, "02", "01") & "|"
            dsTexto = dsTexto & "6|" & Trim(!RucAux) & "|"
            dsTexto = dsTexto & Trim(Str(gfRedond(!Total, 0))) & "|"
            dsTexto = dsTexto & Trim(!ApePatAux) & "|"
            dsTexto = dsTexto & Trim(!ApeMatAux) & "|"
            If Not IsNull(Trim(!NomAux)) Then
                n_Posicion = InStr(1, Trim(!NomAux), " ")
                If n_Posicion <> 0 Then
                    dsTexto = dsTexto & Left(Trim(!NomAux), n_Posicion - 1) & "|"
                    dsTexto = dsTexto & Mid(Trim(!NomAux), n_Posicion + 1) & "|"
                Else
                    dsTexto = dsTexto & Trim(!NomAux) & "||"
                End If
            Else
                dsTexto = dsTexto & "||"
            End If
            dsTexto = dsTexto & Trim(!RazAux) & "|"
''            Write #1, SUBSTR(dsTexto, 2, Len(dsTexto) - 2)
            Print #1, dsTexto
            'dnContador = dnContador + 1
            pgbEtapa1.Value = dnContador
            .MoveNext
         Loop Until .EOF
      End With
      Exit Do
   Loop
   Close #1
   porstTGEMP.Close
End Sub

Private Sub ppEtapa_02()   ' Generacion de Texto en File Ingresos
   Dim dnContador As Integer, n_Posicion As Integer
   Dim dsTexto, dsFile As String
   dnContador = 0
   PgBEtapa2.Min = 0
   With porstTGEMP
      .Source = "Select RucEmp From TGEMP Where CodEmp='" & gsCodEmp & "'"
      .Open
   End With
   'Open "C:\Owl-paqu\Angel.TXT" For Output As #1
   dsFile = "Ingresos.TXT"
   CmnDlgUbica.FileName = dsFile
   CmnDlgUbica.ShowSave
   Open dsFile For Output As #2
   Do
      With porstCOVtaDoc
         If .RecordCount = 0 Then
            Exit Do
         End If
         .MoveFirst
         PgBEtapa2.Max = .RecordCount
         PgBEtapa2.Value = PgBEtapa2.Min
         Do
            dnContador = dnContador + 1
            dsTexto = Trim(Str(dnContador)) & "|"
            dsTexto = dsTexto & "6|" & porstTGEMP!RUCEmp & "|"
            dsTexto = dsTexto & gsAnoAct & "|"
            dsTexto = dsTexto & IIf(!TpoPer = TPOPER_JUR, "02", "01") & "|"
            dsTexto = dsTexto & "6|" & Trim(!RucAux) & "|"
            dsTexto = dsTexto & Trim(Str(gfRedond(!Total, 0))) & "|"
            dsTexto = dsTexto & Trim(!ApePatAux) & "|"
            dsTexto = dsTexto & Trim(!ApeMatAux) & "|"
            'dsTexto = dsTexto & Trim(!NomAux) & "|"
            If Not IsNull(Trim(!NomAux)) Then
                n_Posicion = InStr(1, Trim(!NomAux), " ")
                If n_Posicion <> 0 Then
                    dsTexto = dsTexto & Left(Trim(!NomAux), n_Posicion - 1) & "|"
                    dsTexto = dsTexto & Mid(Trim(!NomAux), n_Posicion + 1) & "|"
                Else
                    dsTexto = dsTexto & Trim(!NomAux) & "||"
                End If
            Else
                dsTexto = dsTexto & "||"
            End If
            dsTexto = dsTexto & Trim(!RazAux) & "|"
''            Write #2, SUBSTR(dsTexto, 2, Len(dsTexto) - 2)
            Print #2, dsTexto
            'dnContador = dnContador + 1
            PgBEtapa2.Value = dnContador
            .MoveNext
         Loop Until .EOF
      End With
      Exit Do
   Loop
   Close #2
   porstTGEMP.Close
End Sub

