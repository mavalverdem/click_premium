VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOutl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fExporTRegistro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5640
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   9375
   Icon            =   "expotregistro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5644.487
   ScaleMode       =   0  'User
   ScaleWidth      =   9375
   Begin Threed.SSFrame sfmProgreso 
      Height          =   465
      Left            =   75
      TabIndex        =   14
      Top             =   5760
      Width           =   9170
      _Version        =   65536
      _ExtentX        =   16175
      _ExtentY        =   820
      _StockProps     =   14
      Caption         =   " Procesando archivo : "
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      ShadowStyle     =   1
      Begin MSComctlLib.ProgressBar pgbProgreso 
         Height          =   225
         Left            =   60
         TabIndex        =   15
         Top             =   210
         Width           =   9080
         _ExtentX        =   16007
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   5025
      Left            =   75
      TabIndex        =   10
      Top             =   555
      Width           =   9170
      _ExtentX        =   16166
      _ExtentY        =   8864
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
      TabCaption(0)   =   "Exportación"
      TabPicture(0)   =   "expotregistro.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmCuadro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmCuadro(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin Threed.SSFrame frmCuadro 
         Height          =   4530
         Index           =   1
         Left            =   6530
         TabIndex        =   0
         Top             =   60
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   7990
         _StockProps     =   14
         Caption         =   " Parametros "
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
         Begin VB.ComboBox cboPeriodo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            ItemData        =   "expotregistro.frx":0028
            Left            =   140
            List            =   "expotregistro.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2625
            Width           =   2240
         End
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Index           =   0
            Left            =   140
            TabIndex        =   3
            Top             =   795
            Width           =   2240
         End
         Begin VB.DriveListBox drbUnidad 
            Height          =   315
            Index           =   0
            Left            =   140
            TabIndex        =   2
            Top             =   495
            Width           =   2240
         End
         Begin VB.Label lblDato 
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   140
            TabIndex        =   4
            Top             =   2385
            Width           =   900
         End
         Begin VB.Label lblDato 
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   140
            TabIndex        =   1
            Top             =   250
            Width           =   1005
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   4530
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   60
         Width           =   6335
         _Version        =   65536
         _ExtentX        =   11174
         _ExtentY        =   7990
         _StockProps     =   14
         Caption         =   " Información a Procesar "
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
         Font3D          =   1
         ShadowStyle     =   1
         Begin MSOutl.Outline outParametro 
            Height          =   3585
            Left            =   255
            TabIndex        =   12
            Top             =   330
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   6324
            _StockProps     =   77
            ForeColor       =   16711680
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
            MouseIcon       =   "expotregistro.frx":002C
            PicturePlus     =   "expotregistro.frx":0048
            PictureMinus    =   "expotregistro.frx":0142
            PictureLeaf     =   "expotregistro.frx":023C
            PictureOpen     =   "expotregistro.frx":075E
            PictureClosed   =   "expotregistro.frx":0858
         End
         Begin Threed.SSCheck chkNuevo 
            Height          =   255
            Left            =   255
            TabIndex        =   13
            Top             =   4140
            Width           =   2940
            _Version        =   65536
            _ExtentX        =   5186
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Nueva Información Trabajador"
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
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   3810
            Index           =   1
            Left            =   45
            Shape           =   4  'Rounded Rectangle
            Top             =   225
            Width           =   6255
         End
         Begin VB.Image imgGrafico 
            Height          =   240
            Left            =   5775
            Top             =   150
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9375
      _Version        =   65536
      _ExtentX        =   16536
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
         Left            =   8435
         TabIndex        =   8
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
         Picture         =   "expotregistro.frx":0D7A
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   8045
         TabIndex        =   9
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "expotregistro.frx":0D96
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   390
         TabIndex        =   7
         Top             =   120
         Width           =   7265
      End
   End
End
Attribute VB_Name = "fExporTRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_Index As Integer                              ' Indice para bucle
Private s_Registro As String                            ' Codigo del registro
'[
Private Function pfGenera_Archivo(ByVal s_Archivo As String, ByVal s_Mensage As String, ByVal s_Sentencia As String) As Boolean
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, sCaracter As String, sExpresion As String
  Dim nRegistro As Long, nRegistros As Long, nColumnas As Long
  Dim porstSeleccion As ADODB.Recordset
  
  On Error GoTo Error
  
  ' Selecciono información de proceso
  Set porstSeleccion = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sentencia)
  sCaracter = "|"
  If Not (porstSeleccion.BOF And porstSeleccion.EOF) Then
    ' Creo objeto de archivo
    Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
    Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
    ' Inicializo variables
    nRegistros = porstSeleccion.RecordCount: nRegistro = 0
    nColumnas = porstSeleccion.Fields.Count
    ' Inicializo progreso
    pgbProgreso.Value = pgbProgreso.Min
    pgbProgreso.Max = nRegistros
    pgbProgreso.Value = pgbProgreso.Min
    sfmProgreso.Caption = " Procesando Información: " & s_Mensage & " - " & Right(s_Archivo, 18) & " "
    While Not porstSeleccion.EOF
      psRegistro = pfParrafo_Texto(porstSeleccion, nColumnas, sCaracter)
      potxtFileExp.WriteLine psRegistro
      ' Incremento porcentaje de progreso
      nRegistro = nRegistro + 1
      pgbProgreso.Value = nRegistro
      DoEvents
      porstSeleccion.MoveNext
    Wend
    ' Cierro objeto y saco de memoria
    potxtFileExp.Close
    porstSeleccion.Close
  End If
  pfGenera_Archivo = True
  GoTo Finalizar

Error:
  MsgBox "Error: " & Err.Number & " : " & Err.Description, vbInformation
  
Finalizar:
  Set porstSeleccion = Nothing
  Set potxtFileExp = Nothing
  Set pofsoFileExp = Nothing

End Function
Private Function pfParrafo_Texto(ByVal o_rstParrafo As ADODB.Recordset, ByVal n_Columnas As Long, s_Caracter As String) As String
  Dim nSecuencia As Long, nBinary As Integer
  Dim nTipoDato As String
  
  pfParrafo_Texto = ""
  For nSecuencia = 0 To (n_Columnas - 1)
    nTipoDato = IIf((o_rstParrafo(nSecuencia).Type = adSmallInt Or o_rstParrafo(nSecuencia).Type = adInteger Or o_rstParrafo(nSecuencia).Type = adDouble Or o_rstParrafo(nSecuencia).Type = adCurrency Or o_rstParrafo(nSecuencia).Type = adNumeric), TipoDato.Numero, IIf(o_rstParrafo(nSecuencia).Type = adChar Or o_rstParrafo(nSecuencia).Type = adVarChar, TipoDato.Caracter, IIf(o_rstParrafo(nSecuencia).Type = adDBDate, TipoDato.FECHA, IIf(o_rstParrafo(nSecuencia).Type = adDBTimeStamp, TipoDato.Caracter, TipoDato.Caracter))))
    nBinary = o_rstParrafo(nSecuencia).Type
    If (nTipoDato = TipoDato.FECHA Or nBinary = adDBTimeStamp) Then
      pfParrafo_Texto = pfParrafo_Texto & gdl_Funcion.aTexto(IIf(IsNull(o_rstParrafo(nSecuencia).Value), "", Format(o_rstParrafo(nSecuencia).Value, IIf(nBinary = adDBTimeStamp, s_FmtFeHoMysql_0, s_FormatoFecha)))) & s_Caracter
    Else
      pfParrafo_Texto = pfParrafo_Texto & gdl_Funcion.SacaEntRetApos(IIf((nBinary = adLongVarBinary Or IsNull(o_rstParrafo(nSecuencia).Value)), "", o_rstParrafo(nSecuencia).Value)) & s_Caracter
    End If
  Next nSecuencia

End Function
Private Function pfParametro_Seleccion(ByVal n_Elemento As Integer, ByRef sArchivo As String) As String
  Dim sExpresion As String, sCondicionExt As String
  
  pfParametro_Seleccion = ""
  sArchivo = ps_Anyo & Left(cboPeriodo.Text, 2)
  Select Case n_Elemento
   Case 2      ' Establecimientos Propios del Empleador
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT cdgepr, indepr FROM plestablecimientopropio WHERE estadoepr='" & s_Estado_Act & "'"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".esp"
   Case 3      ' Empleadores a Quienes Destaco o Desplazo Personal
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT codeqd, acteqd, fechaini_eqd, fechafin_eqd FROM plempresasqdes WHERE estadoeqd='" & s_Estado_Act & "'"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".edd"
   Case 4      ' Empleadores que me Destacan o Desplazan Personal
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT codqmd, actqmd, fechaini_qmd, fechafin_qmd FROM plempresasqmdes WHERE estadoqmd='" & s_Estado_Act & "'"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".med"
   Case 5      ' Datos Personales Trabajador, Pensionista, Personal en Formación - Modalidad Formativa Laboral u Otros y Personal de Terceros
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, Null AS paisemisor, psn.fecnacimiento, psn.apepaterno, psn.apematerno, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.nombres, (CASE WHEN psn.sexopsn='" & s_Estado_Ina & "' THEN '" & s_Estado_Act & "' ELSE '" & s_Estado_Blq & "' END) AS sexopsn, psn.nacionalidad, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.codldn AS prefijo_fono, LEFT(psn.telefono, 9) AS telefono, LEFT(psn.correoelect, 50) AS correoelect, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.codvia END) AS codvia, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.nomviadirec END) AS nomviadirec, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.numerdirec END) AS numerdirec, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.intedirec END) AS numerodpto, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS interior, Null AS manzana, Null AS lote, Null AS kilometro, Null AS block, Null AS etapa, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.codzona END) AS codzona, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.nomzondirec END) AS nomzondirec, "
    pfParametro_Seleccion = pfParametro_Seleccion & "LEFT(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.refedirec END, 40) AS refedirec, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN Null ELSE psn.ubigeodir END) AS ubigeodir, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS codvia2, Null AS nomviadirec2, Null AS numerdirec2, Null AS numerodpto2, Null AS interior2, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS manzana2, Null AS lote2, Null AS kilometro2, Null AS block2, Null AS etapa2, Null AS codzona2, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS nomzondirec2, Null AS refedirec2, Null AS ubigeodir2, '" & s_Estado_Act & "' AS essaluddirec "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".ide"
   Case 6      ' Datos del Trabajador - Tipo planilla 01:02:06
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN chkrl='" & s_Estado_Ina & "' THEN '01' ELSE '02' END) AS regimenlabor, est.grado, psn.codpfs, psn.chkdis, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.numeroafp, (CASE WHEN IFNULL(psn.chksctrp, '0')='" & s_Estado_Ina & "' THEN Null ELSE psn.chksctrp END) AS chksctrp, "
    pfParametro_Seleccion = pfParametro_Seleccion & "IFNULL(con.tipcon, '99') AS tipcon, psn.chkreg, psn.chkmax, psn.chknoc, psn.afilsindical, psn.periodicidad, res.importe_mn, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.estadopsn='I' THEN '" & s_Estado_Ina & "' ELSE '" & s_Estado_Act & "' END) AS situacion, psn.chkqui, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE psn.cgoconfianza WHEN '" & s_Estado_Act & "' THEN '" & s_Estado_Blq & "' WHEN '" & s_Estado_Blq & "' THEN '" & s_Estado_Act & "' ELSE '" & s_Estado_Ina & "' END) AS cgoconfianza, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.tippago, psn.cmbcatocupacional, IFNULL(psn.cmbtributacion, '" & s_Estado_Ina & "') AS tibutacion, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS casnroruc "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('01', '02', '06') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plparametroafp cfg ON cfg.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn AND res.pdoano=cfg.pdoano AND res.codcpc=cfg.cpcbasico "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "LEFT JOIN plestudios est ON psn.codcls=est.codcls AND psn.codpsn=est.codpsn AND est.grado = (SELECT MAX(e.grado) FROM plestudios e WHERE e.codcls=psn.codcls AND e.codpsn=psn.codpsn) "
    pfParametro_Seleccion = pfParametro_Seleccion & "LEFT JOIN plcontrato con ON psn.codcls=con.codcls AND psn.codpsn=con.codpsn AND con.estadocon='" & s_Estado_Act & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".tra"
   Case 7      ' Datos del Pensionista - Tipo planilla 04
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.codtpt, LEFT(afp.codsunat, 2), psn.numeroafp, psn.tippago "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('04') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plentidadafp afp ON psn.codafp=afp.codafp "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".pen"
   Case 8      ' Datos del Personal en Formación - Modalidad Formativa Laboral u Otros - Tipo planilla 03
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, Null AS paisemisor, '03' AS modalidad, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.segmedico='" & s_Estado_Ina & "' THEN '" & s_Estado_Act & "' ELSE '" & s_Estado_Blq & "' END) AS seguromed, "
    pfParametro_Seleccion = pfParametro_Seleccion & "est.grado, psn.codpfs, psn.resfamiliar, psn.chkdis, psn.forprofesional, psn.chknoc "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('03') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "LEFT JOIN plestudios est ON psn.codcls=est.codcls AND psn.codpsn=est.codpsn AND est.grado = (SELECT MAX(e.grado) FROM plestudios e WHERE e.codcls=psn.codcls AND e.codpsn=psn.codpsn) "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".pfl"
   Case 9      ' Datos del Personal de Terceros - Tipo planilla 05
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, Null AS paisemisor, pte.codest, pte.sctrp "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('05') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plterceros pte ON res.codcls=pte.codcls AND res.codpsn=pte.codpsn AND res.pdoano=pte.ano  AND res.pdomes=pte.mes "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".ter"
   Case 10     ' Datos Períodos
    For n_Index = 1 To 5
      ' inicio o reinicio registro
      sExpresion = Choose(n_Index, "Null", "psn.codtpt", "(CASE WHEN IFNULL(psn.codeps, '99')='99' THEN '00' ELSE '01' END)", "LEFT(afp.codsunat, 2)", "(CASE WHEN psn.cobsctr='" & s_Estado_Ina & "' THEN Null ELSE psn.cobsctr END)")
      pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, Null AS paisemisor, "
      pfParametro_Seleccion = pfParametro_Seleccion & "(CASE LEFT(cls.tipo, 2) WHEN '03' THEN '5' WHEN '04' THEN '2' WHEN '05' THEN '4' ELSE '1' END) AS categoria, "
      pfParametro_Seleccion = pfParametro_Seleccion & "'" & n_Index & "' AS tipo_registro, "
      pfParametro_Seleccion = pfParametro_Seleccion & Choose(n_Index, "dxr.fecingreso", "dxr.fecingreso", "dxr.fecingreso", "psn.fecingregpen", "dxr.fecingreso") & " AS facha_inicio, "
      pfParametro_Seleccion = pfParametro_Seleccion & "Null AS fecha_final, " & sExpresion & " AS motivo_registro, Null AS Servicio "
      pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
      pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND dxr.estadopsn<>'I' "
      pfParametro_Seleccion = pfParametro_Seleccion & IIf(n_Index <> 5, "", "AND psn.cobsctr<>'" & s_Estado_Ina & "' ")
      ' filtrar nuevos y clase planilla
      If chkNuevo.Value Then
        pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' "
      End If
      'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
      pfParametro_Seleccion = pfParametro_Seleccion & "UNION ALL "
    Next n_Index
    ' fin registro
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE LEFT(cls.tipo, 2) WHEN '03' THEN '5' WHEN '04' THEN '2' WHEN '05' THEN '4' ELSE '1' END) AS categoria, "
    pfParametro_Seleccion = pfParametro_Seleccion & "'" & s_Estado_Act & "' AS tipo_registro, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS facha_inicio, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.fecbaja AS fecha_final, psn.finperiodo AS motivo_registro, Null AS Servicio "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND dxr.estadopsn='I' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, numdociden, tipo_registro"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".per"
   Case 11     ' Datos de los Establecimientos Donde Labora el Trabajador
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "est.ruc, est.codest "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plestalaboral est ON res.codcls=est.codcls AND res.codpsn=est.codpsn AND res.pdoano=est.ano AND res.pdomes=est.mes "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".est"
   Case 12     ' Lugar de Formación de Personal en Formación - Modalidad Formativa Laboral u Otros y de Destaque del Personal de Terceros
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE LEFT(cls.tipo, 2) WHEN '03' THEN '5' ELSE '4' END) AS categoria, "
    pfParametro_Seleccion = pfParametro_Seleccion & "est.codest "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plpersonal psn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('03', '05') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plestalaboral est ON res.codcls=est.codcls AND res.codpsn=est.codpsn AND res.pdoano=est.ano AND res.pdomes=est.mes "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    ' filtrar nuevos y clase planilla
    If chkNuevo.Value Then
      pfParametro_Seleccion = pfParametro_Seleccion & "AND DATE_FORMAT(dxr.fecingreso, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND dxr.estadopsn='A' "
    End If
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "RP_" & ps_RucEmpresa & ".lug"
   Case 14     ' Datos Derechohabientes - Altas
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, LEFT(dcf.codsunat, 2), fam.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "fam.fecnacimiento, fam.apepaterno, fam.apematerno, fam.nombres, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN fam.sexofam='" & s_Estado_Ina & "' THEN '" & s_Estado_Act & "' ELSE '" & s_Estado_Blq & "' END) AS sexofam, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN fam.vinculo='" & s_Estado_Ina & "' THEN '" & s_Estado_Act & "' ELSE fam.vinculo END) AS vinculo, "
    pfParametro_Seleccion = pfParametro_Seleccion & "fam.tipdocpaternidad, fam.acrepaternidad, Null AS mes_concepcion, "
    pfParametro_Seleccion = pfParametro_Seleccion & "fam.codvia, fam.nomviadom, fam.numerdom, fam.intedom, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS interior, Null AS manzana, Null AS lote, Null AS kilometro, Null AS block, Null AS etapa, "
    pfParametro_Seleccion = pfParametro_Seleccion & "fam.codzona, fam.nomzonadom, fam.refedom, fam.ubigeodom, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS codvia2, Null AS nomviadirec2, Null AS numerdirec2, Null AS numerodpto2, Null AS interior2, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS manzana2, Null AS lote2, Null AS kilometro2, Null AS block2, Null AS etapa2, Null AS codzona2, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS nomzondirec2, Null AS refedirec2, Null AS ubigeodir2, '" & s_Estado_Act & "' AS essaluddirec, "
    pfParametro_Seleccion = pfParametro_Seleccion & "Null AS prefijo_fono, Null AS telefono, Null AS correoelect "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plfamiliares fam "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON psn.codcls=fam.codcls AND psn.codpsn=fam.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dcf ON fam.coddci=dcf.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND year(fam.fecalta)='" & ps_Anyo & "' AND month(fam.fecalta)='" & Left(cboPeriodo.Text, 2) & "' "
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = Format(gdl_Funcion.NumeroDiasMes(Left(cboPeriodo.Text, 2), ps_Anyo), "00") & Left(cboPeriodo.Text, 2) & ps_Anyo
    sArchivo = "RD_" & ps_RucEmpresa & "_" & sArchivo & "_ALTA.TXT"
   Case 15     ' Datos Derechohabientes - Bajas
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2), psn.numdociden, LEFT(dcf.codsunat, 2), fam.numdociden, Null AS paisemisor, "
    pfParametro_Seleccion = pfParametro_Seleccion & "fam.fecnacimiento, fam.apepaterno, fam.apematerno, fam.nombres, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN fam.vinculo='" & s_Estado_Ina & "' THEN '" & s_Estado_Act & "' ELSE fam.vinculo END) AS vinculo, "
    pfParametro_Seleccion = pfParametro_Seleccion & "fam.fecbaja, fam.motivoina "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plfamiliares fam "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON psn.codcls=fam.codcls AND psn.codpsn=fam.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plresultado res ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dcf ON fam.coddci=dcf.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND fam.estadofam='" & s_Estado_Ina & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND year(fam.fecalta)='" & ps_Anyo & "' AND month(fam.fecalta)='" & Left(cboPeriodo.Text, 2) & "' "
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = Format(gdl_Funcion.NumeroDiasMes(Left(cboPeriodo.Text, 2), ps_Anyo), "00") & Left(cboPeriodo.Text, 2) & ps_Anyo
    sArchivo = "RD_" & ps_RucEmpresa & "_" & sArchivo & "_BAJA.TXT"
   Case 17     ' Prestadores Servicio Rentas Cuarta Categoria
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.apepaterno, psn.apematerno, psn.nombres, "
    pfParametro_Seleccion = pfParametro_Seleccion & "(CASE WHEN psn.naciextrapsn = '" & s_Estado_Ina & "' THEN '" & s_Estado_Blq & "' ELSE '" & s_Estado_Act & "' END) AS domiciliadopsn, "
    pfParametro_Seleccion = pfParametro_Seleccion & "psn.cmbtributacion "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plcomprobantect ser "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON ser.codcls=psn.codcls AND ser.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE year(ser.fecemision)='" & ps_Anyo & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND month(ser.fecemision)='" & Left(cboPeriodo.Text, 2) & "' "
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY psn.apepaterno, psn.apematerno, psn.nombres"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".ps4"
   Case 18     ' Trabajador - Otras Rentas quinta categoría
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, emp.ruc AS rucempleador, "
    pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mn "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plresultado res "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plempleadores emp ON res.codcls=emp.codcls AND res.codpsn=emp.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plcfgempresa cfg ON res.pdoano=cfg.pdoano "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON psn.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND res.codcpc IN (cfg.remempordin, cfg.remempextra) "
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY res.codcls, LEFT(dci.codsunat, 2), psn.numdociden "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, psn.numdociden"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".or5"
   Case 19     ' Trabajador - Datos jornada laboral
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, "
    pfParametro_Seleccion = pfParametro_Seleccion & "FORMAT(TRUNCATE(SUM(asi.horanormal), 1), 0) AS horordin, "
    pfParametro_Seleccion = pfParametro_Seleccion & "FORMAT(((SUM(asi.horanormal)-TRUNCATE(SUM(asi.horanormal), 0))*60), 0) AS minordin, "
    pfParametro_Seleccion = pfParametro_Seleccion & "FORMAT(TRUNCATE(SUM(asi.horatipo1+asi.horatipo2+asi.horatipo3), 1), 0) AS horsobre, "
    pfParametro_Seleccion = pfParametro_Seleccion & "FORMAT(((SUM(asi.horatipo1+asi.horatipo2+asi.horatipo3)-TRUNCATE(SUM(asi.horatipo1+asi.horatipo2+asi.horatipo3), 0))*60), 0) AS minsobre "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plasistencia asi "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON asi.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('01', '02') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE pdo.anopdo='" & ps_Anyo & "' AND pdo.mespdo='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND ((DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND psn.estadopsn<>'I') "
    pfParametro_Seleccion = pfParametro_Seleccion & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND psn.estadopsn='I')) "
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY LEFT(dci.codsunat, 2), psn.numdociden "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, psn.numdociden"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".jor"
   Case 20     ' Trabajador - Dias subsidiados y otros no laborados
    For n_Index = 1 To 6
      sExpresion = Choose(n_Index, "enfer", "natal", "licen", "accid", "vacacion", "falta")
      sCondicionExt = Choose(n_Index, "'21'", "'22','28'", "'05', '26'", "'20'", "'23'", "'07'")
      
      pfParametro_Seleccion = pfParametro_Seleccion & "SELECT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, " & IIf(n_Index = 6, sCondicionExt, "asi.codmdi_" & Left(sExpresion, 5)) & " AS codmdi, "
      pfParametro_Seleccion = pfParametro_Seleccion & "asi." & Choose(n_Index, "enfermedad", "diaprepostnatal", "licencia", "accidente", "diavacaciones", "diafalta") & " AS dia_suspension "
      pfParametro_Seleccion = pfParametro_Seleccion & "FROM plasistencia asi "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
      pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      pfParametro_Seleccion = pfParametro_Seleccion & "WHERE ((DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND psn.estadopsn<>'I') "
      pfParametro_Seleccion = pfParametro_Seleccion & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND psn.estadopsn='I')) "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND pdo.anopdo='" & ps_Anyo & "' "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND pdo.mespdo='" & Left(cboPeriodo.Text, 2) & "' "
      If n_Index <= 5 Then
        pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(asi.codmdi_" & Left(sExpresion, 5) & ", '') IN (" & sCondicionExt & ") "
        pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(asi.fechaini" & IIf(n_Index = 5, "", "_") & sExpresion & ", '')<>'' "
        pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(asi.fechafin" & IIf(n_Index = 5, "", "_") & sExpresion & ", '')<>'' "
      End If
      pfParametro_Seleccion = pfParametro_Seleccion & "AND asi." & Choose(n_Index, "enfermedad", "diaprepostnatal", "licencia", "accidente", "diavacaciones", "diafalta") & ">=1 "
      pfParametro_Seleccion = pfParametro_Seleccion & IIf(n_Index = 6, "ORDER BY dcisunat, numdociden", "UNION ALL ")
    Next n_Index
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".snl"
   Case 21     ' Trabajador - Detalle ingresos, tributos y descuentos - Tipo planilla 01:02
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, cpc.codsunat, "
    pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnd, "
    pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnp "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plresultado res "
    pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plclasplan cls ON res.codcls=cls.codcls "
    pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plconceplanilla cpc ON res.codcls=cpc.codcls AND res.codcpc=cpc.codcpc "
    pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND LEFT(cls.tipo, 2) IN('01', '02') "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(cpc.codsunat,'') NOT IN('', '0100', '0200', '0300', '0400', '0500', '0600','0700', '0800', '1000', '2000') "
    pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY res.codcls, LEFT(dci.codsunat, 2), psn.numdociden, cpc.codsunat "
    
    ' Incluye proceso de CTS mayo y noviembre
    If (Left(cboPeriodo.Text, 2) = "05" Or Left(cboPeriodo.Text, 2) = "11") Then
      pfParametro_Seleccion = pfParametro_Seleccion & "UNION "
      pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, cpc.codsunat, "
      pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnd, "
      pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnp "
      pfParametro_Seleccion = pfParametro_Seleccion & "FROM plctsresultado res "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plclasplan cls ON res.codcls=cls.codcls "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plconceplanilla cpc ON res.codcls=cpc.codcls AND res.codcpc=cpc.codcpc "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' AND res.pdomes='" & Format((CLng(Left(cboPeriodo.Text, 2)) - 1), "00") & "' "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND LEFT(cls.tipo, 2) IN('01', '02') "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(cpc.codsunat,'') NOT IN('', '0100', '0200', '0300', '0400', '0500', '0600','0700', '0800', '1000', '2000') "
      pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY res.codcls, LEFT(dci.codsunat, 2), psn.numdociden, cpc.codsunat "
    End If
    
    ' Comision y seguro de AFP sin importe
    For n_Index = 1 To 2
      sExpresion = Choose(n_Index, "cpcporcen", "cpcseguro")
    
      pfParametro_Seleccion = pfParametro_Seleccion & "UNION "
      pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, cpc.codsunat, "
      pfParametro_Seleccion = pfParametro_Seleccion & "0.00 AS importe_mnd, 0.00 AS importe_mnp "
      pfParametro_Seleccion = pfParametro_Seleccion & "FROM pldatoresultado dxr "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plclasplan cls ON dxr.codcls=cls.codcls "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plpersonal psn ON dxr.codcls=psn.codcls AND dxr.codpsn=psn.codpsn "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plperiodo pdo ON dxr.codcls=pdo.codcls AND dxr.codpdo=pdo.codpdo "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plparametroafp prm ON pdo.anopdo=prm.pdoano "
      pfParametro_Seleccion = pfParametro_Seleccion & "JOIN plconceplanilla cpc ON cls.codcls=cpc.codcls AND prm." & sExpresion & "=cpc.codcpc "
      pfParametro_Seleccion = pfParametro_Seleccion & "WHERE ((dxr.codafp NOT IN ('00','99') "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND pdo.anopdo='" & ps_Anyo & "' AND pdo.mespdo='" & Left(cboPeriodo.Text, 2) & "' "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND LEFT(cls.tipo, 2) IN('01', '02') "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(cpc.codsunat,'') NOT IN('', '0100', '0200', '0300', '0400', '0500', '0600','0700', '0800', '1000', '2000')) "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND NOT EXISTS(SELECT * "
      pfParametro_Seleccion = pfParametro_Seleccion & "FROM plresultado res "
      pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.codcls= dxr.codcls "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND res.codpsn= dxr.codpsn "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND res.codcpc= cpc.codcpc "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdoano=pdo.anopdo "
      pfParametro_Seleccion = pfParametro_Seleccion & "AND res.pdomes=pdo.mespdo "
      pfParametro_Seleccion = pfParametro_Seleccion & ")) "
    Next n_Index
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, numdociden, codsunat, importe_mnd"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".rem"
   Case 22     ' Pensionista - Detalle ingresos, tributos y descuentos - Tipo planilla 04
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2)  AS dcisunat, psn.numdociden, cpc.codsunat, "
    pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnd, "
    pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnp "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plresultado res "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plconceplanilla cpc ON res.codcls=cpc.codcls AND res.codcpc=cpc.codcpc "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON res.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('04') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND cpc.codsunat NOT IN ('0100', '0200', '0300', '0400', '0500', '0600', '0601', '0602', '0604', '0605', '0606', '0607', '0608', '0609', '0610', '0612', '0613', '0614', '0700', '0702', '0704', '0705', '0800', '0801', '0802', '0803', '0804', '0805', '0806', '0807', '0808', '0810', '0811') "
    'pfParametro_Seleccion = pfParametro_Seleccion & "AND " & sCondicion & " "
    pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY res.codcls, LEFT(dci.codsunat, 2), psn.numdociden, cpc.codsunat "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, psn.numdociden"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".pen"
   Case 24     ' Personal en Formación - Modalidad Formativa Laboral y Otros: Monto Pagado
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2)  AS dcisunat, psn.numdociden, "
    pfParametro_Seleccion = pfParametro_Seleccion & "ROUND(SUM(res.importe_mn), 2) AS importe_mnd "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plresultado res "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plconceplanilla cpc ON res.codcls=cpc.codcls AND res.codcpc=cpc.codcpc "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON res.codcls=cls.codcls AND LEFT(cls.tipo, 2) = '03' "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE res.pdoano='" & ps_Anyo & "' AND res.pdomes='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND IFNULL(cpc.codsunat, '') <> '' "
    pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY res.codcls, LEFT(dci.codsunat, 2), psn.numdociden, cpc.codsunat "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, psn.numdociden"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".for"
   Case 26     ' Trabajador - Otras condiciones
    pfParametro_Seleccion = pfParametro_Seleccion & "SELECT DISTINCT LEFT(dci.codsunat, 2) AS dcisunat, psn.numdociden, "
    pfParametro_Seleccion = pfParametro_Seleccion & "'0' AseguraPension, '0' VidaSegAcciden, '' FDereSociArtista, (psn.naciextrapsn +1) Domiciliado "
    pfParametro_Seleccion = pfParametro_Seleccion & "FROM plasistencia asi "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN plclasplan cls ON asi.codcls=cls.codcls AND LEFT(cls.tipo, 2) IN ('01', '02') "
    pfParametro_Seleccion = pfParametro_Seleccion & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
    pfParametro_Seleccion = pfParametro_Seleccion & "WHERE pdo.anopdo='" & ps_Anyo & "' AND pdo.mespdo='" & Left(cboPeriodo.Text, 2) & "' "
    pfParametro_Seleccion = pfParametro_Seleccion & "AND ((DATE_FORMAT(psn.fecingreso, '%Y%m')<='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND psn.estadopsn<>'I') "
    pfParametro_Seleccion = pfParametro_Seleccion & "OR (DATE_FORMAT(psn.fecbaja, '%Y%m')='" & ps_Anyo & Left(cboPeriodo.Text, 2) & "' AND psn.estadopsn='I')) "
    pfParametro_Seleccion = pfParametro_Seleccion & "GROUP BY LEFT(dci.codsunat, 2), psn.numdociden "
    pfParametro_Seleccion = pfParametro_Seleccion & "ORDER BY dcisunat, psn.numdociden"
    
    sArchivo = "0601" & sArchivo & ps_RucEmpresa & ".toc"
  End Select

End Function
']
Private Sub cmdAction_Click(Index As Integer)
  Dim s_OldMessage As String, s_Message As String
  Dim nElementos As Integer, nElemento As Integer
  Dim sExpresion As String, sSeleccion As String

  ' Verifico que existan registros
  If Index = 1 Then Unload Me: Exit Sub
  If cboPeriodo.Text = "" Then Beep: MsgBox "Debe seleccionar el Periodo de Información", vbCritical: cboPeriodo.SetFocus: Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Procesar la " & lblTitle.Caption & " ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    fMenu.panPercent.Visible = True
    Me.Height = 6790

    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    
    ' Genero información seleccionada
    nElementos = 1: nElemento = 0
    While outParametro.ListCount > nElementos
      ' Nivel de proceso o selección de información
      If outParametro.Indent(nElementos) <> 1 Then
        ' Verifico se encuentra seleccionado
        If outParametro.PictureType(nElementos) = outClosed Then
          sSeleccion = pfParametro_Seleccion(nElementos, sExpresion)
          If sSeleccion <> "" Then
            pfGenera_Archivo dlbDirectorio(0).path & "\" & sExpresion, Trim(outParametro.List(nElementos)), sSeleccion
          End If
        End If
        nElemento = nElemento + 1
      End If
      nElementos = nElementos + 1
      fMenu.panPercent.FloodPercent = ((nElementos * 100) \ outParametro.ListCount)
    Wend
  End If
  MsgBox "Finalizo exitosamente el Proceso de " & lblTitle, vbInformation
  GoTo Finalizar

Finalizar:
  ' Reinicializo los mensajes
  Me.Height = 6140
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
    
End Sub
Private Sub drbUnidad_Change(Index As Integer)
  dlbDirectorio(Index).path = drbUnidad(Index).drive
  dlbDirectorio(Index).Refresh
End Sub

Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim sExpresion As String

  'Establece posición y titulo del formulario
  Me.Height = 6140: Me.Width = 9460
  Me.Left = 1080: Me.Top = 80
  
  ' Titulo del formulario y panel
  s_TitleWindow = Me.Caption
  lblTitle = "Información del sistema"
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(2, 2)
  ' Icono y título del formulario
  aElemento(2, 1) = "proceso": aElemento(2, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 1
    aElemento(n_Index, 1) = Choose(n_Index + 1, "genarchi", "cancelar")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Procesar ", "Cancelar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  cmdAction(1).Cancel = True
  
  ' periodos
  For n_Index = 1 To 12: cboPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  
  drbUnidad(0).drive = ps_PathSystem
  dlbDirectorio(0).path = ps_PathSystem
 
  ' Configuro los graficos de objeto
  imgGrafico.Picture = LoadPicture()
  sExpresion = gdl_Procedure.ps_PathImagen & "database.bmp"
  If dir(sExpresion, vbNormal) <> "" Then imgGrafico.Picture = LoadPicture(sExpresion)
  imgGrafico.Refresh
  outParametro.PictureLeaf = imgGrafico
  
  sExpresion = gdl_Procedure.ps_PathImagen & "nocheck.bmp"
  If dir(sExpresion, vbNormal) <> "" Then imgGrafico.Picture = LoadPicture(sExpresion)
  imgGrafico.Refresh
  outParametro.PictureOpen = imgGrafico
  
  sExpresion = gdl_Procedure.ps_PathImagen & "check.bmp"
  If dir(sExpresion, vbNormal) <> "" Then imgGrafico.Picture = LoadPicture(sExpresion)
  imgGrafico.Refresh
  outParametro.PictureClosed = imgGrafico
  
  ' expander - contraer
  'sExpresion = gdl_Procedure.ps_PathImagen & "mencheck.bmp"
  'If dir(sExpresion, vbNormal) <> "" Then imgGrafico.Picture = LoadPicture(sExpresion)
  'imgGrafico.Refresh
  'outParametro.PicturePlus = imgGrafico
  'outParametro.PictureMinus = imgGrafico
  
  ' Configuro el objeto de parametro
  outParametro.Style = outPlusPictureText
  outParametro.AddItem " Exportación de Información"
  outParametro.Indent(0) = 0
  outParametro.PictureType(0) = outLeaf
  
  ' informacion trabajador
  outParametro.ListIndex = -1
  outParametro.AddItem " T-Registro Información Personal"
  outParametro.ListIndex = outParametro.ListCount - 1
  For n_Index = 1 To 11
    outParametro.AddItem Choose(n_Index, " Establecimientos Propios del Empleador", " Empleadores a Quienes Destaco o Desplazo Personal", " Empleadores que me Destacan o Desplazan Personal", " Datos Personales Trabajador, Pensionista, Personal en Formación - Modalidad Formativa Laboral u Otros y Personal de Terceros", " Datos del Trabajador", " Datos del Pensionista", " Datos del Personal en Formación - Modalidad Formativa Laboral u Otros", " Datos del Personal de Terceros", " Períodos", " Datos de los Establecimientos Donde Labora el Trabajador", " Lugar de Formación de Personal en Formación - Modalidad Formativa Laboral u Otros y de Destaque del Personal de Terceros")
  Next n_Index
  ' derechohabiente
  outParametro.ListIndex = -1
  outParametro.AddItem " T-Registro Información DerechoHabiente"
  outParametro.ListIndex = outParametro.ListCount - 1
  For n_Index = 1 To 2
    outParametro.AddItem Choose(n_Index, " Datos de Derechohabientes - ALTAS", " Baja de Derechohabientes - BAJAS")
  Next n_Index
  ' planilla mensual
  outParametro.ListIndex = -1
  outParametro.AddItem " Planilla Mensual - PLAME"
  outParametro.ListIndex = outParametro.ListCount - 1
  For n_Index = 1 To 12
    outParametro.AddItem Choose(n_Index, " Prestador de Servicios con Rentas de Cuarta categoría", " Trabajador: Otras Rentas de Quinta categoría", " Trabajador: Datos de la Jornada Laboral", " Trabajador: Días Subsidiados y Otros no Laborados", " Trabajador: Detalle de los Ingresos, Tributos y Descuentos", " Pensionista: Detalle de los Ingresos, Tributos y Descuentos", " Prestador de Servicios con Rentas de Cuarta Categoría: Detalle de Comprobantes", " Personal en Formación - Modalidad Formativa Laboral y Otros: Monto Pagado", " Personal de Terceros - SCTR ESSALUD", " Trabajador: Tasas SCTR-EsSalud y/o Convenio IES", " Trabajador: Otras Condiciones", " Pensionista: Otras Condiciones")
  Next n_Index
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  fMenu.cmbejercicio.Enabled = True
End Sub
Private Sub outParametro_PictureClick(ListIndex As Integer)
  Dim nPicture As Integer
  ' Actualizo la selección de parametro
  If outParametro.Indent(ListIndex) > 1 Then
    nPicture = outParametro.PictureType(ListIndex)
    outParametro.PictureType(ListIndex) = Choose(nPicture + 1, outOpen, outClosed)
  End If
End Sub
Private Sub outParametro_PictureDblClick(ListIndex As Integer)
  Dim nPicture As Integer
  ' Actualizo la selección y niveles
  If outParametro.Indent(ListIndex) = 1 Then
    outParametro.ListIndex = ListIndex
    nPicture = outParametro.PictureType(ListIndex)
    outParametro.PictureType(ListIndex) = Choose(nPicture + 1, outOpen, outClosed)
    n_Index = ListIndex + 1
    Do While outParametro.Indent(n_Index) <> 1
      outParametro.PictureType(n_Index) = Choose(nPicture + 1, outOpen, outClosed)
      n_Index = n_Index + 1
      If n_Index = outParametro.ListCount Then Exit Do
    Loop
  End If
End Sub
