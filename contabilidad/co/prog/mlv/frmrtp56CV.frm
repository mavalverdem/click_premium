VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmrpt56CV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Compras Ventas TXT"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkanno 
      Alignment       =   1  'Right Justify
      Caption         =   "Emitidos Año Anterior"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cboTpoSem 
      Height          =   315
      ItemData        =   "frmrtp56CV.frx":0000
      Left            =   5760
      List            =   "frmrtp56CV.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1150
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Vista Preliminar"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1320
      Picture         =   "frmrtp56CV.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reporte de validación"
      Top             =   1320
      Width           =   1150
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmrtp56CV.frx":0536
      Left            =   5760
      List            =   "frmrtp56CV.frx":0538
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CmnDlgUbica 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      FileName        =   "formulario3323"
      Filter          =   "txt"
   End
   Begin VB.Label Procesando 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Semestre:"
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
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda:"
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
      Index           =   0
      Left            =   4800
      TabIndex        =   3
      Top             =   45
      Width           =   765
   End
End
Attribute VB_Name = "frmrpt56CV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As ADODB.Connection
Dim contador1 As Integer
Dim contador2 As Integer
Dim sql As String
Dim semestre As String
Dim cadena As String
Dim reporte As ADODB.Recordset
Private Sub cmdAceptar_Click(Index As Integer)

Dim Rstdatos As ADODB.Recordset
Set Rstdatos = New ADODB.Recordset
Dim Rstvalores As ADODB.Recordset
Set Rstvalores = New ADODB.Recordset

If cboTpoSem.ListIndex = 0 Then
    semestre = "('00','01','02','03','04','05','06')"
Else
    semestre = "('07','08','09','10','11','12','13')"
End If

sql = "DROP TABLE IF EXISTS tmpvalores "
cnn.Execute sql

sql = "DROP TABLE IF EXISTS tmpauxiliares "
cnn.Execute sql

sql = "CREATE TABLE tmpvalores ("
sql = sql & "D1 varchar(15) not null, "
sql = sql & "D2 varchar(15) not null, "
sql = sql & "D3 varchar(15) not null, "
sql = sql & "D4 varchar(15) not null, "
sql = sql & "D5 varchar(15) not null, "
sql = sql & "D6 varchar(15) not null, "
sql = sql & "D7 varchar(15) not null, "
sql = sql & "D8 varchar(15) not null, "
sql = sql & "D9 varchar(15) not null, "
sql = sql & "D10 varchar(15) not null, "
sql = sql & "D11 varchar(15) not null, "
sql = sql & "D12 varchar(15) not null, "
sql = sql & "D13 varchar(15) not null, "
sql = sql & "D14 varchar(100) not null) "

cnn.Execute sql

sql = "CREATE TABLE tmpauxiliares "

If chkanno.Value = 0 Then

sql = sql & " select 'Compras' as D1,C1.codaux as D2,razaux as D3,"
'sql = sql & " case C1.codtdc when '07' then 0 else sum(C1.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") end as D4,"
sql = sql & " (select sum(C2.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ")  from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc <> '07' and year(C2.feedoc) ='" & gsAnoAct & "') as D4,"
sql = sql & " (select count(*) from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc <> '07' and year(C2.feedoc) ='" & gsAnoAct & "') as D5,"
'sql = sql & " case C1.codtdc when '07' then sum(C1.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") else 0 end as D6,"
sql = sql & " (select sum(C2.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ")  from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc = '07' and year(C2.feedoc) ='" & gsAnoAct & "') as D6,"
sql = sql & " (select count(*) from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc = '07' and year(C2.feedoc) ='" & gsAnoAct & "') as D7,0 as D8,0 as D9,0 as D10,0 AS D11"
sql = sql & " from cocprdoc  c1"
sql = sql & " inner join tgaux on C1.codemp=tgaux.codemp and C1.codaux=tgaux.codaux"
sql = sql & " where C1.codemp='" & gsCodEmp & "' and C1.pdoano='" & gsAnoAct & "'"
sql = sql & " and C1.mespvs in " & semestre & " and year(C1.feedoc) ='" & gsAnoAct & "' group by  C1.codaux"
sql = sql & " Union All"
sql = sql & " select 'Ventas' as D1,C1.codaux as D2,razaux as D3,"
sql = sql & " 0 as D4,0 as D5,0 as D6,0 as D7,"
'sql = sql & " case C1.codtdc when '07' then 0 else sum(C1.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") end AS D8,"
sql = sql & " (select sum(C2.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") from covtadoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc <> '07' ) as D8,"
sql = sql & " (select count(*) from covtadoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc <> '07' ) as D9,"
'sql = sql & " case C1.codtdc when '07' then sum(C1.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") else 0 end as D10,"
sql = sql & " (select sum(C2.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") from covtadoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc = '07' ) as D10,"
sql = sql & " (select count(*) from covtadoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc = '07' ) as D11"
sql = sql & " from covtadoc  c1"
sql = sql & " inner join tgaux on C1.codemp=tgaux.codemp and C1.codaux=tgaux.codaux"
sql = sql & " where C1.codemp='" & gsCodEmp & "' and C1.pdoano='" & gsAnoAct & "'"
sql = sql & " and C1.mespvs in " & semestre & " group by  C1.codaux"

Else

sql = sql & " select 'Compras' as D1,C1.codaux as D2,razaux as D3,"
'sql = sql & " case C1.codtdc when '07' then 0 else sum(C1.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") end as D4,"
sql = sql & " (select sum(C2.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ")  from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc <> '07' and year(C2.feedoc) ='" & gsAnoAct - 1 & "') as D4,"
sql = sql & " (select count(*) from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc <> '07' and year(C2.feedoc) ='" & gsAnoAct - 1 & "') as D5,"
'sql = sql & " case C1.codtdc when '07' then sum(C1.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ") else 0 end as D6,"
sql = sql & " (select sum(C2.impigv_" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & ")  from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc = '07' and year(C2.feedoc) ='" & gsAnoAct - 1 & "') as D6,"
sql = sql & " (select count(*) from cocprdoc C2  where C2.codemp='" & gsCodEmp & "' and C2.pdoano='" & gsAnoAct & "' and C2.mespvs in " & semestre
sql = sql & " and C2.codaux=C1.codaux and C2.codtdc = '07' and year(C2.feedoc) ='" & gsAnoAct - 1 & "') as D7,0 as D8,0 as D9,0 as D10,0 AS D11"
sql = sql & " from cocprdoc  c1"
sql = sql & " inner join tgaux on C1.codemp=tgaux.codemp and C1.codaux=tgaux.codaux"
sql = sql & " where C1.codemp='" & gsCodEmp & "' and C1.pdoano='" & gsAnoAct & "'"
sql = sql & " and C1.mespvs in " & semestre & " and year(C1.feedoc) ='" & gsAnoAct - 1 & "' group by  C1.codaux"


End If

cnn.Execute sql

Select Case Index
Case 0

'sql = "select SUBSTRING(d2,1,8),SUBSTRING(d2,10,1),sum(d4),'0' AS C1,sum(d5),sum(d6),sum(d7),sum(d8),'0' AS C2,sum(d9),sum(d10),sum(d11),'' AS C3,D3 from tmpauxiliares group by d2"
'sql = "select d2,SUBSTRING(d2,10,1),round(sum(d4)),'0' AS C1,round(sum(d5)),round(sum(d6)),round(sum(d7)),round(sum(d8)),'0' AS C2,round(sum(d9)),round(sum(d10)),round(sum(d11)),'' AS C3,D3 from tmpauxiliares group by d2"
sql = "select SUBSTRING(d2,1,8),SUBSTRING(d2,10,1),round(sum(d4)),'0' AS C1,round(sum(d5)),round(sum(d6)),round(sum(d7)),round(sum(d8)),'0' AS C2,round(sum(d9)),round(sum(d10)),round(sum(d11)),'' AS C3,D3 from tmpauxiliares group by d2 "

Rstdatos.Open sql, cnn, adOpenStatic, adLockOptimistic

If Rstdatos.RecordCount = 0 Then
Else
    Rstdatos.MoveFirst
        For contador1 = 0 To Rstdatos.RecordCount - 1
            
            If (Rstdatos.Fields(2).Value + Rstdatos.Fields(9).Value) > 0 Then

            cadena = "'" & Rstdatos.Fields(0).Value & "','" & Rstdatos.Fields(1).Value & "','" & Rstdatos.Fields(2).Value & "',"
            cadena = cadena & "'" & Rstdatos.Fields(3).Value & "','" & IIf(Rstdatos.Fields(2).Value = 0, 0, Rstdatos.Fields(4).Value) & "','" & Rstdatos.Fields(5).Value & "',"
            cadena = cadena & "'" & Rstdatos.Fields(6).Value & "','" & Rstdatos.Fields(7).Value & "','" & Rstdatos.Fields(8).Value & "',"
            cadena = cadena & "'" & Rstdatos.Fields(9).Value & "','" & IIf(Rstdatos.Fields(8).Value = 0, 0, Rstdatos.Fields(10).Value) & "','" & Rstdatos.Fields(11).Value & "',"
            cadena = cadena & "'" & Rstdatos.Fields(12).Value & "','" & Replace(Rstdatos.Fields(13).Value, "'", "") & "'"
           
            Procesando.Caption = contador1 + 1 & " de " & Rstdatos.RecordCount & " Auxiliar --> " & Rstdatos.Fields(13).Value
            Procesando.Refresh
            
            sql = "Insert Into tmpvalores values( " & cadena & ")"
            cnn.Execute sql
            
            End If
        
        Rstdatos.MoveNext
        Next
End If
Rstdatos.Close

    Dim Rst As ADODB.Recordset
    Dim R As Boolean
    Dim Ruta As String
    
    Set Rst = New ADODB.Recordset
    
    sql = "select D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13 from tmpvalores"
   
    Rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    
    If Rst.RecordCount = 0 Then
    MsgBox " No Existen datos para Generar el Archivo ", vbInformation
    Exit Sub
    End If
    
    CmnDlgUbica.ShowSave
    
    Ruta = CmnDlgUbica.FileName
    
    R = Recordset_a_Csv(Rst, Ruta)
    
    MsgBox " Se generó el archivo " & Ruta & " correctamente en base a " & Rst.RecordCount & " Registros", vbInformation
    If Not Rst.State = adStateOpen Then
        Rst.Close
    End If
    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If

Case 1

Set reporte = New ADODB.Recordset

   With reporte
      .ActiveConnection = cnn
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
     
   With reporte
   If .State = adStateOpen Then .Close
        '.Source = " select SUBSTRING(d2,1,8),SUBSTRING(d2,10,1),sum(d4),'0' AS C1,sum(d5),sum(d6),sum(d7),sum(d8),'0' AS C2,sum(d9),sum(d10),sum(d11),d2,d3 from tmpauxiliares group by d2 "
        .Source = " select SUBSTRING(d2,1,8),SUBSTRING(d2,10,1),round(sum(d4)),'0' AS C1,round(sum(d5)),round(sum(d6)),round(sum(d7)),round(sum(d8)),'0' AS C2,round(sum(d9)),round(sum(d10)),round(sum(d11)),'' AS C3,D3," & IIf(chkanno.Value = 0, gsAnoAct, gsAnoAct - 1) & "  from tmpauxiliares group by d2 "
        .Open
   End With
      
   gpEncabezadoRpt frmMain.rptMain, "Listado Compras Ventas", Date, True, False, reporte
   
   With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptcv.rpt"
      '.MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
   End With

End Select

sql = "DROP TABLE IF EXISTS tmpvalores"
cnn.Execute sql

sql = "DROP TABLE IF EXISTS tmpauxiliares"
cnn.Execute sql

End Sub

Private Sub Form_Load()

With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
End With
cboTpoMon.ListIndex = TPOMON_NAC_IND
    
With cboTpoSem
    .AddItem "1 Semestre", 0
    .AddItem "2 Semestre", 1
End With
cboTpoSem.ListIndex = 0

Set cnn = New ADODB.Connection
If ps_Puerto = "" Then
    cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";connection="
Else
    cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";Port=" & ps_Puerto & ";connection="
End If
cnn.CursorLocation = adUseClient
cnn.Open

End Sub
Function Recordset_a_Csv(rs As Recordset, path As String) As Boolean
    On Error GoTo Err_function
    Dim columna
    Dim fila As Integer
    ' Crea el archivo
    Open path For Output As #1
    ' Se mueve al primer registro
    rs.MoveFirst
    ' recorre todo el recordset
    For fila = 0 To rs.RecordCount - 1
        ' nombre del campo
        Print #1, Trim(rs.Fields(0));
        ' recorre todos los campos
        For columna = 1 To rs.Fields.Count - 1
            ' imprime la fila actual en el fichero
            Print #1, ";" & Trim(rs.Fields(columna));
        Next
            ' escribe una línea en blanco
        Print ""
            ' salto de carro
        Print #1, "" & Chr(13) & Chr(10);
            ' mueve el recordset al siguiente registro
        rs.MoveNext
    Next
    ' cierra el archivo
    Close #1
    Exit Function
Err_function:
    MsgBox Err.Description, vbCritical
    Close
End Function

