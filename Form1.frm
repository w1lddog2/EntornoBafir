VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entorno Bafir - Entorno de Gestion de planta"
   ClientHeight    =   2385
   ClientLeft      =   2370
   ClientTop       =   3180
   ClientWidth     =   11280
   ClipControls    =   0   'False
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":273A
   ScaleHeight     =   2385
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   23
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agrupar ensayos en tabla ensayos"
      Height          =   615
      Left            =   2400
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10200
      TabIndex        =   18
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar a Excel"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   1440
   End
   Begin VB.TextBox text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10200
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":2B58
      Left            =   240
      List            =   "Form1.frx":2B5A
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   975
      Left            =   2880
      TabIndex        =   16
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1720
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label Label13 
      Caption         =   "Densidad Práctica"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Revisión"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Temperatura"
      Height          =   255
      Left            =   10200
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "T2"
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "T90"
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Partida"
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Costo"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Densidad Teorica"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Menu mnuArch 
      Caption         =   "Archivo"
      Begin VB.Menu mnut90 
         Caption         =   "Ingresar T90"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu mnuconsult 
         Caption         =   "Consultas de clientes"
         Begin VB.Menu mnuconsulting 
            Caption         =   "Ingresar consulta"
         End
         Begin VB.Menu mnuconsultbuscar 
            Caption         =   "Buscar Consulta"
         End
         Begin VB.Menu mnuconsultmodcons 
            Caption         =   "Agregar a consulta (modificar)"
         End
         Begin VB.Menu mnuconsultrespond 
            Caption         =   "Responder consulta"
         End
         Begin VB.Menu mnuconsultmodresp 
            Caption         =   "Agregar a respuesta (modificar)"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuDurezas 
         Caption         =   "Durezas y ensayos varios de piezas en producción"
         Begin VB.Menu mnuSolicitar 
            Caption         =   "Solicitar Medicion"
         End
         Begin VB.Menu mnuListado 
            Caption         =   "Buscar Lote"
         End
      End
      Begin VB.Menu mnuHist 
         Caption         =   "Histórico de precios"
      End
      Begin VB.Menu mnuDatos 
         Caption         =   "Datos generales"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuadn 
         Caption         =   "Administración"
         Begin VB.Menu mnuadmform 
            Caption         =   "Ver Formulas de compuestos"
         End
         Begin VB.Menu mnucostoscomp 
            Caption         =   "Costo de compuestos"
         End
         Begin VB.Menu mnuflujfond 
            Caption         =   "Flujo de fondos"
            Begin VB.Menu mnuAltaConcepto 
               Caption         =   "Alta de conceptos"
            End
            Begin VB.Menu mnuAltafondo 
               Caption         =   "Alta de fondos"
            End
            Begin VB.Menu mnuasientopago 
               Caption         =   "Asiento de pago"
            End
            Begin VB.Menu mnuvisflujo 
               Caption         =   "Visualizar flujo de fondos"
            End
         End
      End
      Begin VB.Menu mnuIngenieria 
         Caption         =   "Ingenieria"
         Begin VB.Menu mnuContracc 
            Caption         =   "Espesores (Contracción)"
         End
      End
      Begin VB.Menu mnuLab 
         Caption         =   "Laboratorio"
         Begin VB.Menu mnuAEDseguimiento 
            Caption         =   "AED"
            Begin VB.Menu mnuMezcla 
               Caption         =   "Mezcla"
            End
            Begin VB.Menu mnuBlend 
               Caption         =   "Blend"
            End
            Begin VB.Menu mnuLote 
               Caption         =   "Lote"
            End
         End
         Begin VB.Menu mnuMezc 
            Caption         =   "Aprobación de mezclas"
            Begin VB.Menu mnuAgregarPrima 
               Caption         =   "Agregar Materia Prima"
            End
            Begin VB.Menu mnumezclareg 
               Caption         =   "Registro de Pesada"
            End
            Begin VB.Menu mnuMezclaAprob 
               Caption         =   "Aprobar mezcla"
            End
         End
         Begin VB.Menu mnuDesarrollo 
            Caption         =   "Desarrollo"
            Begin VB.Menu mnuDesarrAdmin 
               Caption         =   "Administrador"
            End
            Begin VB.Menu mnuIngFormula 
               Caption         =   "Ingresar Formula"
            End
            Begin VB.Menu mnuverformula 
               Caption         =   "Ver/Modificar Formula"
            End
         End
         Begin VB.Menu mnuIndi 
            Caption         =   "Indicadores"
            Begin VB.Menu mnuIndicadores 
               Caption         =   "Panel de indicadores"
            End
         End
         Begin VB.Menu mnupeg 
            Caption         =   "Pegassus"
            Begin VB.Menu mnuPegassus 
               Caption         =   "Emitir Informes de lotes viejos (Aplicación modificada)"
            End
         End
         Begin VB.Menu mnuStock 
            Caption         =   "Stock"
         End
         Begin VB.Menu mnuReometroNuevo 
            Caption         =   "Reometro Nuevo"
         End
         Begin VB.Menu mnuReometro 
            Caption         =   "Reómetro Viejo"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuNormaAsign 
            Caption         =   "Asignación de normas a compuestos"
         End
         Begin VB.Menu mnuNormas 
            Caption         =   "Normas"
         End
         Begin VB.Menu mnuCuerdaLab 
            Caption         =   "Cuerdas"
            Begin VB.Menu mnuCuerdaInforme 
               Caption         =   "Informe Realizado"
            End
         End
         Begin VB.Menu mnuLabDur 
            Caption         =   "Ingresar durezas y valores de ensayos de piezas de producción"
         End
         Begin VB.Menu mnuDens 
            Caption         =   "Generar listado de Densidades"
         End
         Begin VB.Menu mnuEnsayos 
            Caption         =   "Ensayos"
            Begin VB.Menu mnuEnsayoAgregarNuevo 
               Caption         =   "Agregar nuevo ensayo"
            End
            Begin VB.Menu mnuEnsayosVer 
               Caption         =   "Ver Ensayos"
            End
            Begin VB.Menu mnuAgregarMedicion 
               Caption         =   "Agregar Medicion"
            End
            Begin VB.Menu mnuAED 
               Caption         =   "AED"
               Begin VB.Menu mnuAEDseg 
                  Caption         =   "Imprimir Seguimiento del Lote"
               End
               Begin VB.Menu mnuAEDnuevo 
                  Caption         =   "Nuevo Ensayo"
               End
               Begin VB.Menu mnuAEDenv 
                  Caption         =   "Ingresar Valores Envejecidos"
               End
               Begin VB.Menu mnuAEDinic 
                  Caption         =   "Buscar Valores iniciales"
               End
               Begin VB.Menu mnuAEDfinal 
                  Caption         =   "Buscar Informes Finales"
               End
            End
            Begin VB.Menu mnucomp 
               Caption         =   "Compresiones"
               Begin VB.Menu mnuCompo 
                  Caption         =   "Ingresar dimensiones originales"
               End
               Begin VB.Menu mnuCompe 
                  Caption         =   "Ingresar dimensiones envejecidas"
               End
               Begin VB.Menu BuscaComp 
                  Caption         =   "Buscar"
               End
            End
            Begin VB.Menu mnuDesgarros 
               Caption         =   "Desgarros"
               Begin VB.Menu mnuDesgarrosDimensiones 
                  Caption         =   "Ingresar dimensiones"
               End
               Begin VB.Menu mnuDesgarrosTraccion 
                  Caption         =   "Ingresar valores de tracción"
               End
               Begin VB.Menu mnuDesgarroBuscar 
                  Caption         =   "Buscar Desgarro"
               End
            End
            Begin VB.Menu mnuAM 
               Caption         =   "Dureza Shore A - Shore M"
            End
            Begin VB.Menu mnuEnsExt 
               Caption         =   "Ensayos externos o varios"
            End
            Begin VB.Menu mnutraccion 
               Caption         =   "Traccion y elongación de probetas"
               Begin VB.Menu mnuDimensiones 
                  Caption         =   "Ingresar valores dimensionales"
               End
               Begin VB.Menu mnuValtracc 
                  Caption         =   "Ingresar valores de tracción modo en grupos"
                  Enabled         =   0   'False
               End
               Begin VB.Menu mnutraccionIndividual 
                  Caption         =   "Ingresar valores de tracción modo individual"
               End
               Begin VB.Menu mnuBuscaValores 
                  Caption         =   "Buscar Valores de tracción"
               End
            End
            Begin VB.Menu mnuFluid 
               Caption         =   "Ensayos de Fluido"
            End
            Begin VB.Menu mnuVisco 
               Caption         =   "Viscosidades"
               Begin VB.Menu mnuingvisc 
                  Caption         =   "Ingresar Nueva Partida"
               End
               Begin VB.Menu mnuingvischist 
                  Caption         =   "Ingreso de Partida Histórica"
               End
               Begin VB.Menu mnuBuscaPart 
                  Caption         =   "Buscar Partida"
               End
               Begin VB.Menu mnuTolerancia 
                  Caption         =   "Tolerancias"
               End
            End
         End
         Begin VB.Menu mnuCotiz 
            Caption         =   "Compuestos para cotización"
            Begin VB.Menu mnubusca 
               Caption         =   "Buscar"
            End
            Begin VB.Menu mnuSol 
               Caption         =   "Solicitar Compuesto"
            End
            Begin VB.Menu mnuRecom 
               Caption         =   "Recomendar Compuesto"
            End
            Begin VB.Menu mnupiezasactivas 
               Caption         =   "Listado de piezas activas"
            End
         End
         Begin VB.Menu mnuNoconfReo 
            Caption         =   "No conformidades de reometro"
            Begin VB.Menu mnuIngNoconfReo 
               Caption         =   "Ingresar"
            End
         End
         Begin VB.Menu mnuPartidasNuevas 
            Caption         =   "Partidas Nuevas de compuesto"
         End
      End
      Begin VB.Menu mnuPCP 
         Caption         =   "Pcp"
         Begin VB.Menu mnuPCPaltapieza 
            Caption         =   "Alta de Pieza"
         End
         Begin VB.Menu mnufd 
            Caption         =   "Ver Formula de desarrollo (FD)"
         End
      End
      Begin VB.Menu mnuProd 
         Caption         =   "Producción"
         Begin VB.Menu mnucargarlote 
            Caption         =   "Cargar lote manual"
         End
         Begin VB.Menu mnuConsumos 
            Caption         =   "Consumos"
         End
         Begin VB.Menu mnucargarlotebuscar 
            Caption         =   "Buscar Lote (reimprimir)"
         End
         Begin VB.Menu mnuMezcladoPend 
            Caption         =   "Mezclado pendiente"
         End
         Begin VB.Menu mnuTiempComun 
            Caption         =   "Tiempos de vulcanizado de piezas Comunes"
         End
         Begin VB.Menu mnuTiempoAed 
            Caption         =   "Tiempos de vulcanizado de piezas AED"
         End
      End
      Begin VB.Menu mnuRec 
         Caption         =   "Recepción"
         Begin VB.Menu mnuDespachar 
            Caption         =   "Despachar Mercaderia"
         End
         Begin VB.Menu mnuverformulakilos 
            Caption         =   "Ver formula de compuesto"
         End
         Begin VB.Menu mnubuscarcuerda 
            Caption         =   "Buscar Cuerda"
         End
         Begin VB.Menu mnuReci 
            Caption         =   "Recibir cuerdas"
         End
         Begin VB.Menu mnuRecComp 
            Caption         =   "Recibir Compuesto"
         End
         Begin VB.Menu mnuMP 
            Caption         =   "Ver Materias primas"
         End
         Begin VB.Menu mnuRecBuscaComp 
            Caption         =   "Buscar Compuesto"
         End
      End
      Begin VB.Menu mnumant 
         Caption         =   "Mantenimiento"
         Begin VB.Menu mnumantpanel 
            Caption         =   "Panel de control"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuDolar 
         Caption         =   "Actualizar Dolar"
      End
      Begin VB.Menu mnuLotesPegassus 
         Caption         =   "Actualizar lotes del pegassus"
      End
      Begin VB.Menu mnumedicionespegasus 
         Caption         =   "Actualizar mediciones en pegassus"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "BackUp de Bafir"
      End
      Begin VB.Menu mnuCrearReg 
         Caption         =   "Crear registo en historico de precios"
      End
      Begin VB.Menu mnuIny 
         Caption         =   "Productividad de Inyectoras"
      End
      Begin VB.Menu mnuTermocupla 
         Caption         =   "Monitorear Termocupla"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpcion 
         Caption         =   "Opciones"
         Begin VB.Menu mnuContra 
            Caption         =   "Contraseñas"
         End
      End
   End
   Begin VB.Menu mnuQuest 
      Caption         =   "?"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual Entorno Bafir"
      End
   End
   Begin VB.Menu mnuCalidad 
      Caption         =   "ISO 9000"
      Begin VB.Menu mnuManualCalidad 
         Caption         =   "Manual de calidad"
      End
      Begin VB.Menu mnuCertificado 
         Caption         =   "Certificado de calidad"
      End
      Begin VB.Menu mnuOrganigrama 
         Caption         =   "Organigrama"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'''''''''del registro de windows
'Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
''''''''del registro de windows
'''''''''para ejecutar el notepad
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1
Private sAppName As String, sAppPath As String
'''''''''para ejecutar el notepad
Public AEDPassword
Public ComputerName As String
Public peRmiso As Integer
Public reg As Integer
Public reg1 As Integer
Public coneXion As Integer
Public lim As Integer
Public logUser As String
Public logPass As String
Public flags As Boolean
Private Sub BuscaComp_Click()
Form1.Enabled = False
Form1.Visible = False
frmBuscaComp.Height = 2745
frmBuscaComp.Text1.Text = ""
frmBuscaComp.Text2.Text = ""
frmBuscaComp.Text3.Text = ""
frmBuscaComp.Command2.Visible = True
frmBuscaComp.Command3.Visible = False
frmBuscaComp.Show
End Sub

Private Sub Command1_Click()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select referencia from traccion group by referencia")
Set rs1 = db.OpenRecordset("Select * from ensayos")

If rs1.RecordCount = 0 Then
    Inicio = 1
Else
    Inicio = rs1.RecordCount + 1
End If

Do Until rs.EOF = True
    rs1.AddNew
    rs1.Fields("Codigo") = Inicio
    rs1.Fields("referencia") = rs.Fields("referencia")
    rs1.Fields("tipo") = "Traccion"
    rs1.Update
    Inicio = Inicio + 1
    rs.MoveNext
Loop

Set rs = db.OpenRecordset("Select tiempo_temperatura from compresion group by tiempo_temperatura")
Set rs1 = db.OpenRecordset("Select * from ensayos")

If rs1.RecordCount = 0 Then
    Inicio = 1
Else
    Inicio = rs1.RecordCount + 1
End If

Do Until rs.EOF = True
    rs1.AddNew
    rs1.Fields("Codigo") = Inicio
    rs1.Fields("referencia") = rs.Fields("tiempo_temperatura")
    rs1.Fields("tipo") = "Compresion"
    rs1.Update
    Inicio = Inicio + 1
    rs.MoveNext
Loop

Set rs = db.OpenRecordset("Select tiemp_temp_oil from fluidos group by tiemp_temp_oil")
Set rs1 = db.OpenRecordset("Select * from ensayos")

If rs1.RecordCount = 0 Then
    Inicio = 1
Else
    Inicio = rs1.RecordCount + 1
End If

Do Until rs.EOF = True
    rs1.AddNew
    rs1.Fields("Codigo") = Inicio
    rs1.Fields("referencia") = rs.Fields("tiemp_temp_oil")
    rs1.Fields("tipo") = "Fluido"
    rs1.Update
    Inicio = Inicio + 1
    rs.MoveNext
Loop

db.Close
End Sub
Private Sub Combo1_Click()
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  Text5.Text = ""
  Text6.Text = ""
If mnut90.Checked <> True Then
      If mnuDatos.Checked = True Then
            If Combo1.Text = "" Then
                r = MsgBox("Debe seleccionar un compuesto", vbCritical + vbOKOnly, "Error")
                Combo1.SetFocus
                Exit Sub
            End If
            Valor = Combo1.Text
 
            'If Combo1.Enabled = True Then
            
            
                Dim db As Database
                Dim rs As Recordset
                Dim strQu As String
                Dim col As Integer
                Dim rs1 As Recordset
                Dim rs2 As Recordset
            
                Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
                If ComputerName = "REOMETRO" Or ComputerName = "MATRIZ3" Or ComputerName = "MATRIZ4" Or ComputerName = "PIRRI" Or ComputerName = "ANY" Then
                
                    Dim cnn As ADODB.Connection
                    Dim rs51 As ADODB.Recordset
                    
                    Set cnn = New ADODB.Connection
                    Set rs51 = New ADODB.Recordset
                    
                    sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
                    
                    With cnn
                         'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
                         .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
                         .Open
                    End With
                    
                    rs51.Open "SELECT REVISION, FECHA, N_FORMULA,DESCRIPCION, COSTO_TOTAL, DENSIDAD,PARTIDA, ESTADO  FROM Formbase WHERE N_FORMULA = '" & Valor & "' ;", cnn, adOpenStatic, adLockReadOnly
                              
                    Set rs1 = db.OpenRecordset("SELECT * FROM tiempos WHERE N_Formula = '" & Valor & "' ;")
                    If rs1.RecordCount = 0 Then
                        Text5.Text = "Consultar"
                        Text6.Text = "Consultar"
                        Text7.Text = "Consultar"
                    Else
                        Text5.Text = rs1.Fields("T90") & ""
                        Text6.Text = rs1.Fields("T2") & ""
                        Text7.Text = Format(rs1.Fields("Temperatura"), "0") & "ºC"
                    End If
                        Text1.Text = (Format(rs51.Fields("Densidad"), "#.###")) & " g/ml"
                        Text2.Text = rs51.Fields("Costo_Total") & " $"
                        On Error Resume Next 'goto tres
                        Text3.Text = CStr(rs51.Fields("Partida"))
        'tres:
                        Text4.Text = rs51.Fields("Estado")
                        If Text4.Text = "0" Then
                            Label9.Caption = "Baja"
                            Label9.ForeColor = &HFF&
                        End If
                        If Text4.Text = "1" Then
                            Label9.Caption = "APROBADO"
                            Label9.ForeColor = &HFF00&
                        End If
                        If Text4.Text = "2" Then
                            Label9.Caption = "ENSAYO"
                            Label9.ForeColor = &HFFFF&
                        End If
                        If Text4.Text = "3" Then
                            Label9.Caption = "OBSERVACION"
                            Label9.ForeColor = &H80FF&
                        End If
                        If Text4.Text = "4" Then
                            Label9.Caption = "DESUSO"
                            Label9.ForeColor = &HFF&
                        End If
                        If Text4.Text = "5" Then
                            Label9.Caption = "DESARROLLO"
                            Label9.ForeColor = &HFFFF&
                        End If
                        If Text4.Text = "6" Then
                            Label9.Caption = "RETENIDA"
                            Label9.ForeColor = &HFF&
                        End If
                        Label7.Caption = rs51.Fields("DESCRIPCION")
                        Label12.Caption = rs51.Fields("revision") & " - " & Format(rs51.Fields("fecha"), "DD/MM/YY")
                        
                         Set rs2 = db.OpenRecordset("SELECT densidad from densidades WHERE compuesto = '" & Valor & "' Order by fecha desc ;")
                        If rs2.RecordCount = 0 Then
                            Text8.Text = "Consultar"
                        Else
                            rs2.MoveFirst
                            Text8.Text = Format(rs2.Fields("densidad"), "0.00")
                        End If
                        
                        
                        cnn.Close
                
                
                
                Else
                 
                    Set rs = db.OpenRecordset("SELECT N_FORMULA,DESCRIPCION, COSTO_TOTAL, DENSIDAD,PARTIDA, ESTADO, REVISION, FECHA  FROM Formbase WHERE N_FORMULA = '" & Valor & "' ;")
                
                
                
                
                Set rs1 = db.OpenRecordset("SELECT * FROM tiempos WHERE N_Formula = '" & Valor & "' ;")
                If rs1.RecordCount = 0 Then
                    Text5.Text = "Consultar"
                    Text6.Text = "Consultar"
                    Text7.Text = "Consultar"
                Else
                    Text5.Text = rs1.Fields("T90")
                    Text6.Text = rs1.Fields("T2")
                    Text7.Text = Format(rs1.Fields("Temperatura"), "0") & "ºC"
                End If
                Text1.Text = (Format(rs.Fields("Densidad"), "#.###")) & " g/ml"
                Text2.Text = rs.Fields("Costo_Total") & " $"
                On Error Resume Next 'goto tres
                Text3.Text = CStr(rs.Fields("Partida"))
'tres:
                Text4.Text = rs.Fields("Estado")
                If Text4.Text = "0" Then
                    Label9.Caption = "Baja"
                    Label9.ForeColor = &HFF&
                End If
                If Text4.Text = "1" Then
                    Label9.Caption = "APROBADO"
                    Label9.ForeColor = &HFF00&
                End If
                If Text4.Text = "2" Then
                    Label9.Caption = "ENSAYO"
                    Label9.ForeColor = &HFFFF&
                End If
                If Text4.Text = "3" Then
                    Label9.Caption = "OBSERVACION"
                    Label9.ForeColor = &H80FF&
                End If
                If Text4.Text = "4" Then
                    Label9.Caption = "DESUSO"
                    Label9.ForeColor = &HFF&
                End If
                If Text4.Text = "5" Then
                    Label9.Caption = "DESARROLLO"
                    Label9.ForeColor = &HFFFF&
                End If
                If Text4.Text = "6" Then
                    Label9.Caption = "RETENIDA"
                    Label9.ForeColor = &HFF&
                End If
                Label7.Caption = rs.Fields("DESCRIPCION")
                Label12.Caption = rs.Fields("revision") & " - " & Format(rs.Fields("fecha"), "DD/MM/YY")
                
                Set rs2 = db.OpenRecordset("SELECT densidad from densidades WHERE compuesto = '" & Valor & "' Order by fecha desc ;")
                If rs2.RecordCount = 0 Then
                    Text8.Text = "Consultar"
                Else
                    rs2.MoveFirst
                    Text8.Text = Format(rs2.Fields("densidad"), "0.00")
                End If
                
                
                
                
                db.Close
            End If
        Else
                If mnuHist.Checked = True Then
                    Valor = Combo1.Text
                    MSFlexGrid1.Clear
                    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
                    Set rs = db.OpenRecordset("HISTORICO_PRECIOS", dbOpenTable)
                    rs.Index = "primarykey"
                    rs.Seek "=", Valor
                    If rs.NoMatch = True Then
                        er = MsgBox("No existe el compuesto", vbCritical + vbOKOnly, "Error")
                        Combo1.SetFocus
                        Exit Sub
                    End If
                    col = rs.Fields.Count
                    compi = rs.Fields(0).Value
                    MSFlexGrid1.Cols = col
                    For col1 = 0 To (col - 1)
                        MSFlexGrid1.TextMatrix(1, col1) = rs.Fields(col1)
                        MSFlexGrid1.TextMatrix(0, col1) = rs.Fields(col1).Name
                    Next
                End If
                db.Close
            End If
    Else
        Dim db1 As Database
        Set db1 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
        Set rs1 = db1.OpenRecordset("Tiempos", dbOpenTable)
        rs1.Index = "primarykey"
        rs1.Seek "=", Combo1.Text
        If rs1.NoMatch = True Then
            er = MsgBox("No existe el compuesto. Desea agregarlo?", vbCritical + vbYesNo, "Error")
            If er = vbYes Then
                rs1.AddNew
                rs1.Fields("N_FORMULA") = Combo1.Text
                rs1.Fields("T90") = 0
                rs1.Fields("T2") = 0
                rs1.Fields("Fecha") = Date
                rs1.Update
                MsgBox ("Ahora seleccione el compuesto y cargue los datos")
                Text5.Enabled = False
                Text6.Enabled = False
                Command1.Enabled = False
                Exit Sub
            Else
                Combo1.SetFocus
                Exit Sub
            End If
        Else
            Text5.Enabled = True
            Text6.Enabled = True
            Command1.Enabled = True
            Text5.Text = ""
            Text6.Text = ""
            Text5.Text = rs1.Fields("T90")
            Text6.Text = rs1.Fields("T2")
            Combo1.Enabled = False
            db1.Close
            If ComputerName = "REOMETRO" Or ComputerName = "MATRIZ3" Or ComputerName = "MATRIZ4" Or ComputerName = "ANY" Then
                cnn.Close
            End If
        End If
    End If
End Sub

Private Sub Command2_Click()
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Text = ""
Text6.Text = ""
Command1.Visible = False
Command2.Visible = False
Combo1.Enabled = True
Combo1.Visible = True
mnut90.Enabled = True
mnut90.Checked = False
mnuDatos.Enabled = True
mnuHist.Enabled = True
mnuCrearReg.Enabled = True
Combo1.Text = ""
Combo1.SetFocus
End Sub

Private Sub Command3_Click()
Form1.Enabled = False
Form1.Visible = False
Form3.Show
End Sub

Private Sub Command4_Click()
decimalAtiempo ("0.72")
End Sub

Private Sub Form_Load()

''''''''''esto compacta la base
'Call DBEngine.CompactDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", "\\Servidor2\e\EntornoBafir\TEMPpartidas de compuesto.mdb", dbLangGeneral, , ";pwd=flanflus")
'DoEvents
'Kill "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
'Name "\\Servidor2\e\EntornoBafir\TEMPpartidas de compuesto.mdb" As "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
''''''''''esto compacta la base
Form2.Show
DoEvents








Static db As Database
Static rs As Recordset
Dim rs1 As Recordset
'''''''''''login''''''''''''''''
registro123 = GetSetting("EBafir", "Valores", "UsadoYa")
ComputerName = UCase(regQuery_A_Key(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName"))
'ComputerName = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows Media\WMSDK\General", "ComputerName")

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from sys where maquina = '" & ComputerName & "'")
Form2.Hide
If rs.Fields("activar") = True Then
    regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "System", "\\Servidor2\e\EntornoBafir\System\sys.exe"
    If Not IsTaskRunning("system") Then
        Shell ("\\Servidor2\e\EntornoBafir\System\sys.exe")
    End If
Else
    regDelete_Sub_Key HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "System"
End If

If ComputerName = "PIRRI" Or ComputerName = "PEDRO" Or ComputerName = "GUSTAVO" Then
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select * from compuestos_nuevos where etiquetado = False and fecha_cierre <> Null")
    faltan = rs.RecordCount
    If faltan <> 0 Then
        Do Until rs.EOF = True
            If Date - rs.Fields("fecha_cierre") >= 2 Then
                etiquetar = etiquetar & " " & rs.Fields("partida") & " ; "
            End If
            rs.MoveNext
        Loop
    sdfsdf = MsgBox("Las siguientes partidas están pendientes de etiquetarse. Por favor, realice el etiquetado de las mismas: " & etiquetar, vbCritical + vbOKOnly, "ATENCION!!!")
    
    End If
    
End If

If ComputerName = "MATRIZ4" Or ComputerName = "MATRIZ3" Or ComputerName = "ANY" Or ComputerName = "GUSTAVO" Then
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select codigo_recomendacion, comp_recomendado from compuestos_para_cotizacion where revisado_ing = False order by codigo_recomendacion")
    If rs.RecordCount <> 0 Then
        textt = "Las siguientes solicitudes de cotización aún no han sido contestadas:"
        Do Until rs.EOF = True
            textt = textt & " " & rs.Fields("codigo_recomendacion") & ";"
            rs.MoveNext
            Loop
        kgl = MsgBox(textt, vbCritical + vbOKOnly, "Solicitudes de cotización")
    End If
End If


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select conexion from loginonline")
rs.MoveLast
If rs.RecordCount = 0 Then
    coneXion = 1
Else
coneXion = rs.Fields("conexion") + 1
End If
db.Close

Dim db20 As Database
Dim rs20 As Recordset

Set db20 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs20 = db20.OpenRecordset("Select * from REG where funcion = 'Time'")

'If (Date) >= CDate(rs20.Fields("dato")) Then
'    asddddd = MsgBox("Fatal Error", vbCritical + vbOKOnly, "Error")
'    End
'End If






'todo esto es el login, no lo borres que es para poner cuando funque para que no se compile ahora
If registro123 = "" Then
    renoi = MsgBox("Esta es la primera vez que utiliza el programa luego de la renovación. Por favor, complete el siguiente formulario para poder realizar en breve el ingreso por nombre de usuario y contraseña", vbInformation + vbOKOnly, "Aviso")
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select Numero From Loginusuarios")
    If rs.RecordCount = 0 Then
        frmSignUp.Text8.Text = "1"
    Else
    rs.MoveLast
       
    frmSignUp.Text8.Text = rs.RecordCount + 1
    End If
    Set rs = db.OpenRecordset("select sectores from loginsectores")
    rs.MoveFirst
    frmSignUp.Combo1.Clear
    Do Until rs.EOF = True
    frmSignUp.Combo1.AddItem (rs.Fields("sectores"))
    rs.MoveNext
    Loop
    frmSignUp.Text4.Text = Date
    db.Close
    frmSignUp.Show (1)
    
End If


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select conexion, Fechaaccion, maquina, Usuario, accion from loginonline")
    rs.AddNew
    rs.Fields("fechaaccion") = Now()
    rs.Fields("maquina") = ComputerName
    rs.Fields("Usuario") = GetSetting("EBafir", "Valores", "Usuario")
    rs.Fields("Accion") = "Sign Up"
    rs.Fields("conexion") = coneXion
    rs.Update
    db.Close
frmLogin.Label1.Caption = "Hola " & GetSetting("EBafir", "Valores", "Nombre") & " por favor ingresá tu contraseña."
frmLogin.Text1.Text = GetSetting("Ebafir", "Valores", "usuario")
''''''''''este es el login
'frmLogin.Show (1)
''''''''''este es el login
usuario = frmLogin.Text1.Text
frmSplash.Show (0)
sendInfo

 '''''''''''''''Permisos
'Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
'Set rs1 = db.OpenRecordset("Select email, permiso from loginusuarios where email = '" & usuario & "'")
'Set rs = db.OpenRecordset("Select * from loginpermisos where tipo = '" & rs1.Fields("permiso") & "'")
    
'mnuLab.Enabled = rs.Fields("lab_completo")
'mnuDurezas.Enabled = rs.Fields("dureza_prod")
'mnuHist.Enabled = rs.Fields("historico_precios")
'mnucomp.Enabled = rs.Fields("Lab_compresion")
'mnuReometro.Enabled = rs.Fields("lab_reometro")
'mnuNoconfReo.Enabled = rs.Fields("lab_noconfreo")
'mnuLabDur.Enabled = rs.Fields("Lab_dureza")
'mnuCuerdaLab.Enabled=rs.Fields("lab_cuerdas")
'mnuTraccion.Enabled = rs.Fields("Lab_traccion")
'mnuFluid.Enabled = rs.Fields("Lab_fluido")
'mnuDens.Enabled = rs.Fields("lab_Densidades")
'mnuVisco.Enabled = rs.Fields("lab_viscosidades")
'mnuCotiz.Enabled = rs.Fields("lab_cotizacion")
'mnuRec.Enabled = rs.Fields("recepcion")
'mnuBack.Enabled = rs.Fields("backup")
'mnuCrearReg.Enabled = rs.Fields("registro_historico_precios")
'mnuIny.Enabled = rs.Fields("inyectoras")
'mnuRecom.Enabled = rs.Fields("lab_cotizacion_recomendar")
'Timer1.Enabled = rs.Fields("contador")
'peRmiso = rs1.Fields("permiso")
'db.Close
'''''''''''''''Permisos

    
    
    

'''''''''''login'''''''''''''''''''

'sendInfo
'frmSplash.Show


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")

Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'servicio'")
If rs.Fields("dato") = "0" Then
    End
End If
'Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'revision'")
'If rs.Fields("dato") <> App.Revision Then
'    sdfsdfsdf = MsgBox("Usted no está usando una versión autorizada de Entorno Bafir. Consulte con el administrador del sistema.", vbCritical + vbOKOnly)
'    End
'End If


''''''''''''''''''este es el login viejo sacar cuando se active login
Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'labkey'")

If ComputerName = "SANDRA" Or ComputerName = "PIRRI" Or ComputerName = "LAB3" Or ComputerName = "REOMETRO" Then
    'contra = InputBox("Ingrese contraseña de Laboratorio", "Contraseña")
    frmPassword.Show (1)
    contra = frmPassword.Password

Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'Formbase_dupl'")
    'If CDate(rs.Fields("dato")) < Date Then
        'If a = "CACA" Then 'CInt(CStr(Date)) > CInt(rs.Fields("dato")) Then
        'Dim rs2 As Recordset
    
        
        'Esto es para hacer un duplicado del formbase, ya que con DAO en XP o 2000 no podia abrir la tabla vinculada de excel
        'Set rs1 = db.OpenRecordset("Select * from formbase")
        'Set rs2 = db.OpenRecordset("Select * from COPIA_FORMBASE")
        'rs2.MoveFirst
        'Do Until rs2.EOF = True
        'rs2.Delete
        'rs2.MoveNext
        'Loop
        '
        'rs1.MoveFirst
        'Do Until rs1.EOF = True
        '    rs2.AddNew
        '    rs2.Fields("N_FORMULA") = rs1.Fields("N_FORMULA")
        '    rs2.Fields("DESCRIPCION") = rs1.Fields("DESCRIPCION")
        '    rs2.Fields("REVISION") = rs1.Fields("REVISION")
        '    rs2.Fields("FECHA") = rs1.Fields("FECHA")
        '    rs2.Fields("DENSIDAD") = rs1.Fields("DENSIDAD")
        '    rs2.Fields("COSTO_TOTAL") = rs1.Fields("COSTO_TOTAL")
        '    rs2.Fields("PARTIDA") = rs1.Fields("PARTIDA")
        '    rs2.Fields("ESTADO") = rs1.Fields("ESTADO")
        '    rs1.MoveNext
        '    rs2.Update
        'Loop
        
    
    'Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'Formbase_dupl'")
    'rs.Edit
    'rs.Fields("dato") = CStr(Date)
    'rs.Update
    'End If
Else
    contra = ""
End If

Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'LabKey'")
If UCase(rs.Fields("dato")) = UCase(contra) Then
'fsdfss1 = MsgBox("Contraseña correcta. Abriendo en modo Laboratorio", vbInformation + vbOKOnly, "Modo Laboratorio")
Form1.Caption = "Entorno Bafir - Entorno de Gestión de planta - Modo Laboratorio"
'Form1.Caption = "Entorno Bafir - Entorno de Gestión de planta"
frmSplash.Label1.Caption = "Modo Laboratorio"
Timer1.Enabled = False

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select codigo_recomendacion, comp_recomendado from compuestos_para_cotizacion where revisado_lab = False order by codigo_recomendacion")
If rs.RecordCount <> 0 Then
    textt = "Las siguientes solicitudes de cotización aún no han sido contestadas:"
    Do Until rs.EOF = True
    textt = textt & " " & rs.Fields("codigo_recomendacion") & ";"
    rs.MoveNext
    Loop
    kgl = MsgBox(textt, vbCritical + vbOKOnly, "Solicitudes de cotización")
End If

frmStock.Controla_minimo
frmConsultas.Controla_consultas


peRmiso = 1
Else
'fsdfss = MsgBox("Contraseña incorrecta. Abriendo en modo normal", vbInformation + vbOKOnly, "Modo Normal")
frmSplash.Label1.Caption = "Modo Normal"
'MsgBox ("Sin mensajes")
Form1.Caption = "Entorno Bafir - Entorno de Gestión de planta - Modo Normal"
frmIndicadores.Command1.Enabled = False
frmIndicadores.Command2.Enabled = False
frmIndicadores.Command3.Enabled = False
mnuBack.Enabled = False
mnuCrearReg.Enabled = False
frmNormas.Command2.Enabled = False
mnuDesgarrosDimensiones.Enabled = False
mnuDesgarrosTraccion.Enabled = False
mnuReometroNuevo.Enabled = False
mnuRecom.Enabled = False
mnuIngNoconfReo.Enabled = False
mnuMezclaAprob.Enabled = False
mnuLabDur.Enabled = False
mnuAEDnuevo.Enabled = False
mnuAEDenv.Enabled = False
mnuCompo.Enabled = False
mnuCompe.Enabled = False
frmEnsExt.Command1.Enabled = False
frmFluidos.Command4.Enabled = False
mnuCuerdaInforme.Enabled = False
mnuDimensiones.Enabled = False
mnuIngFormula.Enabled = False
mnuverformula.Enabled = False
mnuValtracc.Enabled = False
mnutraccionIndividual.Enabled = False
mnuTolerancia.Enabled = False
mnuTermocupla.Enabled = False
mnuingvisc.Enabled = False
mnuingvischist.Enabled = False
mnuconsultrespond.Enabled = False
mnuconsultmodresp.Enabled = False
mnuOpcion.Enabled = False
mnuEnsayoAgregarNuevo.Enabled = False
mnuAEDseguimiento.Enabled = False
frmModPartidasNuevas.Command1.Enabled = False
End If

frmSplash.SetFocus
frmAbout.Visible = False
''''''''''''''''''este es el login viejo sacar cuando se active login
reg = 1000
lim = 120
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Combo1.Clear
MSFlexGrid1.Visible = False

'Completado de tiempos de reometro por parte de la maquina del reometro USA ADO
If ComputerName = "REOMETRO" Then '

    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim dbloc As Database
    Dim rsloc As Recordset '

    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset


    sPathBase = "\\REOMETRO\tisa\reo.mdb" '

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";"
        .Open
    End With

    
    rst.Open "SELECT IDParameter, Compound FROM Parametertable", cnn, adOpenStatic, adLockReadOnly
    
    Set dbloc = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
    Set rsloc = dbloc.OpenRecordset("SELECT * FROM Tiempos")


    rsloc.MoveFirst
    Do Until rsloc.EOF = True
        rsloc.Delete
        rsloc.MoveNext
    Loop
    rst.MoveFirst

    Do Until rst.EOF = True
        rsloc.AddNew
        rsloc.Fields("N_formula") = rst.Fields("Compound")
        rsloc.Fields("IDParameter") = rst.Fields("IDParameter")
        rsloc.Update
        rst.MoveNext
    Loop
    rsloc.MoveFirst
    rst.Close
   Do Until rsloc.EOF = True
        rst.Open "SELECT Operatoracceptance, IDParameter, ts2, T90, TempsupI, TempsupF, TempinfI, TempinfF FROM M1 where IDParameter = " & rsloc.Fields("IDParameter") & " And operatoracceptance = True Order by datemeasure desc", cnn, adOpenStatic, adLockReadOnly
        
        
   
        If rst.EOF = True Then
            rst.Close
            rsloc.MoveNext
        Else
            rsloc.Edit
            rsloc.Fields("T2") = rst.Fields("ts2")
            rsloc.Fields("T90") = rst.Fields("T90")
            rsloc.Fields("Temperatura") = (CDbl(rst.Fields("tempsupi")) + CDbl(rst.Fields("tempsupf")) + CDbl(rst.Fields("tempinfi")) + CDbl(rst.Fields("tempinff"))) / 4
            rsloc.Update
            rsloc.MoveNext
            rst.Close
        End If
    Loop


    
     cnn.Close
     dbloc.Close
End If

Dim HistFecha As Date
Dim db1 As Database
Dim strQu As String
Dim Fila As Integer
Dim restaDias As Integer
Dim contador As Integer

On Error GoTo Erri

If peRmiso = 1 Then
    Set db1 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
    Set rs1 = db1.OpenRecordset("REG", dbOpenTable)
    rs1.Index = "primarykey"
    rs1.Seek "=", "Histfecha"
    HistFecha = CDate(rs1.Fields("Dato"))
    db1.Close
    restaDias = (Date) - (HistFecha)

        If restaDias >= 30 Then
            pepe = MsgBox("El tiempo de actualización del Historico ha caducado. Por favor realice la actualización del mismo", vbCritical + vbOKOnly, "Aviso de caducidad")
        End If
End If
        
        
        
        
        'llenado del combo de compuestos

        
        
        
If ComputerName = "REOMETRO" Or ComputerName = "MATRIZ3" Or ComputerName = "MATRIZ4" Or ComputerName = "ANY" Then
'este if define que maquinas tienen XP o 2000

'esto carga el combo para las maquinas con windows XP o 2000
'Dim cnn As ADODB.Connection
'Dim rst As ADODB.Recordset

Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

With cnn
'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
.Open
End With

rst.Open "SELECT N_FORMULA FROM Formbase", cnn, adOpenStatic, adLockReadOnly

        rst.MoveFirst
        Do Until rst.EOF = True
            Combo1.AddItem (rst.Fields("N_FORMULA"))
            rst.MoveNext
        Loop
Else
        'esto carga el combo para las maquinas con windows 98
        Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
        Set rs = db.OpenRecordset("SELECT N_FORMULA FROM Formbase")

        rs.MoveFirst
        a = rs.Fields("N_formula")

        Do While rs.EOF <> True

            rs.MoveNext
        Loop
        rs.MovePrevious
        Fila = rs.RecordCount
        rs.MoveFirst
        For contador = 1 To Fila
            b = rs.Fields("N_FORMULA").Value
        On Error GoTo fIn
            Combo1.AddItem (rs.Fields("N_FORMULA"))
            rs.MoveNext
        Next
End If
fIn:
'''''''''''''''''''backupear base














'''''''''''esto lo pongo por que en las demás maquinas hasta que yo no habro el programa se cuelga
On Error Resume Next
'''''''''''esto lo pongo por que en las demás maquinas hasta que yo no habro el programa se cuelga

Set rs = db.OpenRecordset("Select funcion, dato from reg where funcion = 'Repair'")
If CDate(rs.Fields("dato")) <= Date Then
    Flag = "compactaaa"
    Form2.Show
    'Form2.ProgressBar1.Visible = True
    'Form2.ProgressBar1.Value = 100
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    fso.CopyFile "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_Partidas de compuesto.mdb"
    fso.CopyFile "\\Servidor2\e\EntornoBafir\centralpesado.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_centralpesado.mdb"
    'fso.CopyFile "\\REOMETRO\tisa\reo.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_reo.mdb"
       
    fso.CopyFile "\\SERVIDOR2\Pegassus\IntelligentWorld\Datos\base.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_basepegassus.mdb"
    '
    fso.CopyFile "\\Servidor2\e\produccion\lotes\LoteProducción.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_loteproducción.mdb"
    fso.CopyFile "\\Servidor2\e\administracion\compras\gestión de compras.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_Gestión de compras.mdb"
    fso.CopyFile "\\Servidor2\e\Laboratorio\compuestos\formulas.xls", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_Formulas.xls"
    fso.CopyFile "\\Servidor2\e\Laboratorio\compuestos\compuestos.xls", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_Compuestos.xls"
    fso.CopyFile "\\Servidor2\e\EntornoBafir\AED.mdb", "\\Servidor2\e\EntornoBafir\baseautorepair\" & (Format(Date, "YY-MM-DD")) & "_AED.mdb"
    
    '
    
    Form2.ProgressBar1.Visible = False
    Form2.Hide
    If Err.Number = 0 Then
        rs.Edit
        rs.Fields("dato") = CDbl(rs.Fields("dato")) + 1
        rs.Update
    End If
    'db.Close
    Call DBEngine.CompactDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", "\\Servidor2\e\EntornoBafir\TEMPpartidas de compuesto.mdb", dbLangGeneral, , ";pwd=flanflus")
    DoEvents
    Kill "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
    Name "\\Servidor2\e\EntornoBafir\TEMPpartidas de compuesto.mdb" As "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"
    Call DBEngine.CompactDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", "\\Servidor2\e\EntornoBafir\TEMPcentralpesado.mdb", dbLangGeneral, , ";pwd=flanflus")
    DoEvents
    'Kill "\\Servidor2\e\EntornoBafir\centralpesado.mdb"
    Name "\\Servidor2\e\EntornoBafir\TEMPcentralpesado.mdb" As "\\Servidor2\e\EntornoBafir\centralpesado.mdb"
    Err.Clear
    'Call DBEngine.CompactDatabase("\\Servidor2\pegassus\intelligentworld\datos\base.mdb", "\\Servidor2\pegassus\intelligentworld\datos\TEMPbase.mdb", dbLangGeneral, , ";")
    'a = Err.Description
    
    'DoEvents
    'b = Err.Number
    'If b <> 3356 Then
    '    Kill "\\Servidor2\pegassus\intelligentworld\datos\base.mdb"
    '    Name "\\Servidor2\pegassus\intelligentworld\datos\TEMPbase.mdb" As "\\Servidor2\pegassus\intelligentworld\datos\base.mdb"
    'End If
    Call DBEngine.CompactDatabase("\\Servidor2\e\EntornoBafir\AED.mdb", "\\Servidor2\e\EntornoBafir\TEMPAED.mdb", dbLangGeneral, , "")
    DoEvents
    'Kill "\\Servidor2\e\EntornoBafir\AED.mdb"
    Name "\\Servidor2\e\EntornoBafir\TEMPAED.mdb" As "\\Servidor2\e\EntornoBafir\AED.mdb"
    
    

End If

'VE LA CONFIGURACION REGIONAL
Call confreg
If confr = "punto" Then
    dasdasd = MsgBox("El programa está diseñado para funcionar con la convención de signos del pais, en la cual se fija al 'punto' como separador decimal, y a la 'coma' como separador de miles. Otra configuración hará que la carga de los valores sea errónea. La unica sección que se ha diseñado para funcionar bajo esta convención no standard para el Pais, es la sección de administración", vbCritical + vbOKOnly, "Atención")
    mnuLab.Enabled = False
    mnuProd.Enabled = False
    mnuRec.Enabled = False
    mnumant.Enabled = False
    mnuDespachar.Enabled = False
End If



















'''''''''''''''''''backupear base
Erri:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Err = True Then
    If Err.Number = 0 Then
        ini = MsgBox("La base de datos esta siendo utilizada por otro usuario. Por favor intente más tarde.", vbCritical + vbOKOnly, "Error")
        End
    End If
End If
'db1.Close
db.Close
frmSplash.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
'estas lineas hacen la limpieza de los datos temporales de fluidos,traccion y compresion. fluido todavía no lo prové...
If peRmiso = 1 Then
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs1 = db.OpenRecordset("Select codigo_ensayo from traccion_dimensiones")
If rs1.RecordCount <> 0 Then
    rs1.MoveLast
    If rs1.RecordCount > 100 Then
        rs1.MoveFirst
        For borrartracc = 1 To 10
         
            Set rs = db.OpenRecordset("Select codigo_ensayo, traccion from traccion where codigo_ensayo = " & rs1.Fields("codigo_ensayo"))
            If rs.Fields("traccion") <> "0" Then
            rs1.Delete
            End If
            rs1.MoveNext
        Next
    End If
End If
Set rs1 = db.OpenRecordset("Select codigo_ensayo from compresion_temporal")
If rs1.RecordCount <> 0 Then
    rs1.MoveLast
    If rs1.RecordCount > 100 Then
        rs1.MoveFirst
        For borrartracc = 1 To 10
            
            Set rs = db.OpenRecordset("Select codigo_ensayo, compresion_porc from compresion where codigo_ensayo = " & rs1.Fields("codigo_ensayo") & "")
            If rs.Fields("compresion_porc") <> "0" Then
            rs1.Delete
            End If
            rs1.MoveNext
        Next
    End If
End If
Set rs1 = db.OpenRecordset("Select codigo from fluido_temp")
If rs1.RecordCount <> 0 Then
    rs1.MoveLast
    If rs1.RecordCount > 50 Then
        rs1.MoveFirst
        For borrartracc = 1 To 10
                       
            Set rs = db.OpenRecordset("Select codigo, var_vol from fluidos where codigo = " & rs1.Fields("codigo"))
            If rs.Fields("var_vol") <> "0" Then
            rs1.Delete
            End If
            rs1.MoveNext
        Next
    End If
End If
db.Close
End If
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select conexion, Fechaaccion, maquina, Usuario, accion from loginonline")
    rs.AddNew
    rs.Fields("fechaaccion") = Now()
    rs.Fields("maquina") = ComputerName
    rs.Fields("Usuario") = GetSetting("EBafir", "Valores", "Usuario")
    rs.Fields("Accion") = "Sign Out"
    rs.Fields("conexion") = Form1.coneXion
    rs.Update
    Do
        rs.MoveLast
        If rs.RecordCount > 100 Then
            rs.MoveFirst
            rs.Delete
        Else
            Exit Do
        End If
    Loop
    db.Close
End
End Sub

Private Sub mnuAcerca_Click()
frmAbout.Show
End Sub

Private Sub mnuadmform_Click()
asdasd = InputBox("Ingrese Clave", "Clave")
If asdasd <> "bigotes" Then
    Exit Sub
End If


Me.Enabled = False
Dim db As Database
Dim rs As Recordset


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_Formula, estado from formbase where estado = 1 or estado = 3 order by N_FORMULA")
'Set rs = db.OpenRecordset("Select N_Formula, estado from copia_formbase where estado = 1 or estado = 3 order by N_FORMULA")
'Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select N_formula from partes_copia group by N_formula")


frmverFormulasP.List1.Clear
frmverFormulasP.Text1.Text = ""
frmverFormulasP.MSFlexGrid1.Clear
frmverFormulasP.Label4.Caption = ""

Do Until rs.EOF = True
    frmverFormulasP.List1.AddItem (rs.Fields("N_formula"))
    rs.MoveNext
Loop


db.Close
frmverFormulasP.Show
frmverFormulasP.List1.SetFocus
End Sub

Private Sub mnuAEDenv_Click()
Dim Valor As String
frmPassword.Show (1)
contr = frmPassword.Password
'contr = InputBox("Ingrese password", "Password")

If contr = "" Then
    Exit Sub
End If

Dim db As Database
Dim rs As Recordset


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDcontra where contraseña = '" & contr & "'")

responsablecon = rs.Fields("nombre")

Ensayo = InputBox("Ingrese nº de ensayo", "Nº de ensayo")

Set rs = db.OpenRecordset("Select * from AEDorig where ensayo = " & Ensayo)
If rs.RecordCount = 0 Then
    asdd = MsgBox("No se encuentra el ensayo", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

Set rs = db.OpenRecordset("Select * from AEDenv where ensayo = " & Ensayo)

If rs.RecordCount <> 0 Then
    sdfff = MsgBox("El ensayo ya se encuentra cerrado", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

asddd = MsgBox("Se le solicitarán algunas mediciones. Ingréselas como se las pida el programa", vbInformation + vbOKOnly, "Atención")


frmAEDvisual.Show (1)

If frmAEDvisual.AEDvisual = 0 Then
    Exit Sub
Else
    visual = frmAEDvisual.AEDvisual
End If

Dim w21 As String
Dim w22 As String
Dim w23 As String
Dim w24 As String

Dim w31 As String
Dim w32 As String
Dim w33 As String
Dim w34 As String

Dim w41 As String
Dim w42 As String
Dim w43 As String
Dim w44 As String

Dim p21 As String
Dim p22 As String
Dim p23 As String

Dim p31 As String
Dim p32 As String
Dim p33 As String

Dim p41 As String
Dim p42 As String
Dim p43 As String

Dim pe2 As String
Dim pe3 As String
Dim pe4 As String

Dim wa2 As String
Dim wa3 As String
Dim wa4 As String

Dim sH2 As String
Dim sH3 As String
Dim sH4 As String

w21 = InputBox("Ingrese espesor W 1/4", "Oring 2")
If w21 = "" Then
    Exit Sub
End If
w21 = punToaComa(w21)
w22 = InputBox("Ingrese espesor W 2/4", "Oring 2")
If w22 = "" Then
    Exit Sub
End If
w22 = punToaComa(w22)
w23 = InputBox("Ingrese espesor W 3/4", "Oring 2")
If w23 = "" Then
    Exit Sub
End If
w23 = punToaComa(w23)
w24 = InputBox("Ingrese espesor W 4/4", "Oring 2")
If w24 = "" Then
    Exit Sub
End If
w24 = punToaComa(w24)

w31 = InputBox("Ingrese espesor W 1/4", "Oring 3")
If w31 = "" Then
    Exit Sub
End If
w31 = punToaComa(w31)
w32 = InputBox("Ingrese espesor W 2/4", "Oring 3")
If w32 = "" Then
    Exit Sub
End If
w32 = punToaComa(w32)
w33 = InputBox("Ingrese espesor W 3/4", "Oring 3")
If w33 = "" Then
    Exit Sub
End If
w33 = punToaComa(w33)
w34 = InputBox("Ingrese espesor W 4/4", "Oring 3")
If w34 = "" Then
    Exit Sub
End If
w34 = punToaComa(w34)

w41 = InputBox("Ingrese espesor W 1/4", "Oring 4")
If w41 = "" Then
    Exit Sub
End If
w41 = punToaComa(w41)
w42 = InputBox("Ingrese espesor W 2/4", "Oring 4")
If w42 = "" Then
    Exit Sub
End If
w42 = punToaComa(w42)
w43 = InputBox("Ingrese espesor W 3/4", "Oring 4")
If w43 = "" Then
    Exit Sub
End If
w43 = punToaComa(w43)
w44 = InputBox("Ingrese espesor W 4/4", "Oring 4")
If w44 = "" Then
    Exit Sub
End If
w44 = punToaComa(w44)

p21 = InputBox("Ingrese perímetro 1/3", "Oring 2")
If p21 = "" Then
    Exit Sub
End If
p21 = punToaComa(p21)
p22 = InputBox("Ingrese perímetro 2/3", "Oring 2")
If p22 = "" Then
    Exit Sub
End If
p22 = punToaComa(p22)
p23 = InputBox("Ingrese perímetro 3/3", "Oring 2")
If p23 = "" Then
    Exit Sub
End If
p23 = punToaComa(p23)

p31 = InputBox("Ingrese perímetro 1/3", "Oring 3")
If p31 = "" Then
    Exit Sub
End If
p31 = punToaComa(p31)
p32 = InputBox("Ingrese perímetro 2/3", "Oring 3")
If p32 = "" Then
    Exit Sub
End If
p32 = punToaComa(p32)
p33 = InputBox("Ingrese perímetro 3/3", "Oring 3")
If p33 = "" Then
    Exit Sub
End If
p33 = punToaComa(p33)

p41 = InputBox("Ingrese perímetro 1/3", "Oring 4")
If p41 = "" Then
    Exit Sub
End If
p41 = punToaComa(p41)
p42 = InputBox("Ingrese perímetro 2/3", "Oring 4")
If p42 = "" Then
    Exit Sub
End If
p42 = punToaComa(p42)
p43 = InputBox("Ingrese perímetro 3/3", "Oring 4")
If p43 = "" Then
    Exit Sub
End If
p43 = punToaComa(p43)

pe2 = InputBox("Ingrese peso", "Oring 2")
If pe2 = "" Then
    Exit Sub
End If
pe2 = punToaComa(pe2)
pe3 = InputBox("Ingrese peso", "Oring 3")
If pe3 = "" Then
    Exit Sub
End If
pe3 = punToaComa(pe3)

pe4 = InputBox("Ingrese peso", "Oring 4")
If pe4 = "" Then
    Exit Sub
End If
pe4 = punToaComa(pe4)

wa21 = InputBox("Ingrese peso en agua 1/3", "Oring 2")
If wa21 = "" Then
    Exit Sub
End If
wa21 = punToaComa(wa21)

wa22 = InputBox("Ingrese peso en agua 2/3", "Oring 2")
If wa22 = "" Then
    Exit Sub
End If
wa22 = punToaComa(wa22)
wa23 = InputBox("Ingrese peso en agua 3/3", "Oring 2")
If wa23 = "" Then
    Exit Sub
End If
wa23 = punToaComa(wa23)

wa2 = (CDbl(wa21) + CDbl(wa22) + CDbl(wa23)) / 3

wa31 = InputBox("Ingrese peso en agua 1/3", "Oring 3")
If wa31 = "" Then
    Exit Sub
End If
wa31 = punToaComa(wa31)
wa32 = InputBox("Ingrese peso en agua 2/3", "Oring 3")
If wa32 = "" Then
    Exit Sub
End If
wa32 = punToaComa(wa32)
wa33 = InputBox("Ingrese peso en agua 3/3", "Oring 3")
If wa33 = "" Then
    Exit Sub
End If
wa33 = punToaComa(wa33)
wa3 = (CDbl(wa31) + CDbl(wa32) + CDbl(wa33)) / 3

wa41 = InputBox("Ingrese peso en agua 1/3", "Oring 4")
If wa41 = "" Then
    Exit Sub
End If
wa41 = punToaComa(wa41)
wa42 = InputBox("Ingrese peso en agua 2/3", "Oring 4")
If wa42 = "" Then
    Exit Sub
End If
wa42 = punToaComa(wa42)
wa43 = InputBox("Ingrese peso en agua 3/3", "Oring 4")
If wa43 = "" Then
    Exit Sub
End If
wa43 = punToaComa(wa43)

wa4 = (CDbl(wa41) + CDbl(wa42) + CDbl(wa43)) / 3



sH2 = 0
For tomar = 1 To 5
sH2 = sH2 + CDbl(punto_por_coma(InputBox("Ingrese la dureza Oring 1", "Ingrese la dureza " & tomar & "/5")))
Next
sH2 = sH2 / 5


'sH2 = InputBox("Ingrese dureza", "Oring 2")
If sH2 = "" Then
    Exit Sub
End If
'sH2 = punToaComa(sH2)

sH3 = 0
For tomar = 1 To 5
sH3 = sH3 + CDbl(punto_por_coma(InputBox("Ingrese la dureza Oring 2", "Ingrese la dureza " & tomar & "/5")))
Next
sH3 = sH3 / 5



'sH3 = InputBox("Ingrese dureza", "Oring 3")
If sH3 = "" Then
    Exit Sub
End If
'sH3 = punToaComa(sH3)


sH4 = 0
For tomar = 1 To 5
sH4 = sH4 + CDbl(punto_por_coma(InputBox("Ingrese la dureza Oring 3", "Ingrese la dureza " & tomar & "/5")))
Next
sH4 = sH4 / 5



'sH4 = InputBox("Ingrese dureza", "Oring 4")
If sH4 = "" Then
    Exit Sub
End If
'sH4 = punToaComa(sH4)








w2 = (CDbl(w21) + CDbl(w22) + CDbl(w23) + CDbl(w24)) / 4
w3 = (CDbl(w31) + CDbl(w32) + CDbl(w33) + CDbl(w34)) / 4
w4 = (CDbl(w41) + CDbl(w42) + CDbl(w43) + CDbl(w44)) / 4
p2 = (CDbl(p21) + CDbl(p22) + CDbl(p23)) / 3
p3 = (CDbl(p31) + CDbl(p32) + CDbl(p33)) / 3
p4 = (CDbl(p31) + CDbl(p32) + CDbl(p33)) / 3
di2 = (CDbl(p2) / 3.14) - (2 * CDbl(w2))
di3 = (CDbl(p3) / 3.14) - (2 * CDbl(w3))
di4 = (CDbl(p4) / 3.14) - (2 * CDbl(w4))
wa2 = (CDbl(wa21) + CDbl(wa22) + CDbl(wa23)) / 3
wa3 = (CDbl(wa31) + CDbl(wa32) + CDbl(wa33)) / 3
wa4 = (CDbl(wa41) + CDbl(wa42) + CDbl(wa43)) / 3

estirar = ((50 / 100) + 1 - (41.04 / (di2 * 3.14))) * ((di2 * 3.14) / 2)
'41.04 es el perimetro de la polea
dfsdfsf = MsgBox("Ahora debe traccionar el oring con las mordazas de 12.9. Estire primero a " & estirar & " mm para el módulo al 50%, y luego hasta la ruptura.", vbInformation + vbOKOnly, "Atención")

al50 = InputBox("Ingrese los kilos que obtuvo al " & estirar & " mm (módulo 50%).", "Módulo 50%")
If al50 = "" Then
    Exit Sub
End If
al50 = punToaComa(al50)

tracc = InputBox("Ingrese los kilos que obtuvo en la ruptura.", "Tracción")
If tracc = "" Then
    Exit Sub
End If
tracc = punToaComa(tracc)

elon = InputBox("Ingrese la elongación en mm a la ruptura.", "Elongación")
If elon = "" Then
    Exit Sub
End If
elon = punToaComa(elon)
evaluaacion = MsgBox("Debe ingresar un veredicto" & Chr(13) & "Aprobado", vbInformation + vbYesNo, "Aprobación")
If evaluaacion = vbYes Then
    aprobado = True
Else
    aprobado = False
End If

'Set rs = db.OpenRecordset("Select * from AEDenv where ensayo")
'foto = MsgBox("Desea agregar fotos?", vbInformation + vbYesNo, "Agregar fotos")
'If foto = vbYes Then
 







'End If

rs.AddNew
rs.Fields("ensayo") = Ensayo
rs.Fields("responsable") = responsablecon
rs.Fields("espesor2") = w2
rs.Fields("espesor3") = w3
rs.Fields("espesor4") = w4
rs.Fields("perimetro2") = p2
rs.Fields("perimetro3") = p3
rs.Fields("perimetro4") = p4
rs.Fields("diamint2") = di2
rs.Fields("diamint3") = di3
rs.Fields("diamint4") = di4
rs.Fields("peso2") = pe2
rs.Fields("peso3") = pe3
rs.Fields("peso4") = pe4
rs.Fields("pagua2") = wa2
rs.Fields("pagua3") = wa3
rs.Fields("pagua4") = wa4
rs.Fields("densidad2") = (0.9971 * CDbl(pe2)) / (CDbl(pe2) - CDbl(wa2))
rs.Fields("densidad3") = (0.9971 * CDbl(pe3)) / (CDbl(pe3) - CDbl(wa3))
rs.Fields("densidad4") = (0.9971 * CDbl(pe4)) / (CDbl(pe4) - CDbl(wa4))
rs.Fields("dureza2") = sH2
rs.Fields("dureza3") = sH3
rs.Fields("dureza4") = sH4
rs.Fields("mpa2") = (CDbl(tracc) / (CDbl(w2) * CDbl(w2) * 1.57)) * 10
rs.Fields("elong2") = (((2 * CDbl(elon)) + 40.5 - (di2 * 3.14)) / (di2 * 3.14)) * 100
rs.Fields("Modulo502") = (CDbl(al50) / (CDbl(w2) * CDbl(w2) * 1.57)) * 10
rs.Fields("Visual") = visual
rs.Fields("fecha") = Date
rs.Fields("aprobado") = aprobado
rs.Update
db.Close
sdfdsf = MsgBox("Se han cargado exitosamente los valores envejecidos", vbInformation + vbOKOnly, "Atención")
End Sub

Private Sub mnuAEDfinal_Click()
Ensayo = InputBox("Ingrese el nº de ensayo que quiere visualizar", "Nº de ensayo")
If Ensayo = "" Then
    Exit Sub
End If

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDorig where ensayo = " & Ensayo)
Set rs1 = db.OpenRecordset("Select * from AEDenv where ensayo = " & Ensayo)
If rs.RecordCount = 0 Then
    dsfsdf = MsgBox("No existe el ensayo", vbCritical + vbOKOnly, "Error")
    Exit Sub
Else
    If rs1.RecordCount = 0 Then
        dsfsdf = MsgBox("No se han ingresado los valores envejecidos", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
End If
wi = (CDbl(rs.Fields("espesor1")) + CDbl(rs.Fields("espesor2")) + CDbl(rs.Fields("espesor3")) + CDbl(rs.Fields("espesor4")) + CDbl(rs.Fields("espesor5"))) / 5
wf = (CDbl(rs1.Fields("espesor2")) + CDbl(rs1.Fields("espesor3")) + CDbl(rs1.Fields("espesor4"))) / 3
di = (CDbl(rs.Fields("diamint1")) + CDbl(rs.Fields("diamint2")) + CDbl(rs.Fields("diamint3")) + CDbl(rs.Fields("diamint4")) + CDbl(rs.Fields("diamint5"))) / 5
df = (CDbl(rs1.Fields("diamint2")) + CDbl(rs1.Fields("diamint3")) + CDbl(rs1.Fields("diamint4"))) / 3
Pesoi = (CDbl(rs.Fields("peso1")) + CDbl(rs.Fields("peso2")) + CDbl(rs.Fields("peso3")) + CDbl(rs.Fields("peso4")) + CDbl(rs.Fields("peso5"))) / 5
pesof = (CDbl(rs1.Fields("peso2")) + CDbl(rs1.Fields("peso3")) + CDbl(rs1.Fields("peso4"))) / 3
densi = (CDbl(rs.Fields("densidad1")) + CDbl(rs.Fields("densidad2")) + CDbl(rs.Fields("densidad3")) + CDbl(rs.Fields("densidad4")) + CDbl(rs.Fields("densidad5"))) / 5
densf = (CDbl(rs1.Fields("densidad2")) + CDbl(rs1.Fields("densidad3")) + CDbl(rs1.Fields("densidad4"))) / 3
shi = (CDbl(rs.Fields("dureza1")) + CDbl(rs.Fields("dureza2")) + CDbl(rs.Fields("dureza3")) + CDbl(rs.Fields("dureza4")) + CDbl(rs.Fields("dureza5"))) / 5
shf = (CDbl(rs1.Fields("dureza2")) + CDbl(rs1.Fields("dureza3")) + CDbl(rs1.Fields("dureza4"))) / 3
tracci = rs.Fields("MPA1")
traccf = rs1.Fields("MPA2")
elongi = rs.Fields("elong1")
elongf = rs1.Fields("elong2")
modi = rs.Fields("modulo501")
modf = rs1.Fields("modulo502")
frmAEDfinal.Text1.Text = Ensayo
frmAEDfinal.Text2.Text = Format(wi, "0.00")
frmAEDfinal.Text3.Text = Format(wf, "0.00")
frmAEDfinal.Text4.Text = Format(di, "0.00")
frmAEDfinal.Text5.Text = Format(df, "0.00")
frmAEDfinal.Text6.Text = Format(Pesoi, "0.00")
frmAEDfinal.Text7.Text = Format(pesof, "0.00")
frmAEDfinal.Text8.Text = Format(densi, "0.00")
frmAEDfinal.Text9.Text = Format(densf, "0.00")
frmAEDfinal.Text10.Text = Format(shi, "0.00")
frmAEDfinal.Text11.Text = Format(shf, "0.00")
frmAEDfinal.Text12.Text = Format(tracci, "0.00")
frmAEDfinal.Text13.Text = Format(traccf, "0.00")
frmAEDfinal.Text14.Text = Format(elongi, "0.00")
frmAEDfinal.Text15.Text = Format(elongf, "0.00")
frmAEDfinal.Text16.Text = Format(modi, "0.00")
frmAEDfinal.Text17.Text = Format(modf, "0.00")

vw = (wf * 100 / wi) - 100
If vw >= 0 Then
vw = "+" & vw
Else
vw = vw
End If

vd = (df * 100 / di) - 100
If vd >= 0 Then
vd = "+" & vd
Else
vd = vd
End If

vp = (pesof * 100 / Pesoi) - 100
If vp >= 0 Then
vp = "+" & vp
Else
vp = vp
End If

vdens = Format((densf * 100 / densi) - 100, "0.00")
If vdens >= 0 Then
vdens = "+" & vdens
Else
vdens = vdens
End If

vsh = (shf * 100 / shi) - 100
If vsh >= 0 Then
vsh = "+" & vsh
Else
vsh = vsh
End If

vtracc = (traccf * 100 / tracci) - 100
If vtracc >= 0 Then
vtracc = "+" & vtracc
Else
vtracc = vtracc
End If

velong = (elongf * 100 / elongi) - 100
If velong >= 0 Then
velong = "+" & velong
Else
velong = velong
End If
If modi = 0 Or modi = "0" Then
    vmod = 100
Else
    vmod = (modf * 100 / modi) - 100
    If vmod >= 0 Then
    vmod = "+" & vmod
    Else
    vmod = vmod
    End If
End If
'frmAEDfinal.Text18.Text = Format(vw, "0.00")
'frmAEDfinal.Text19.Text = Format(vd, "0.00")
'frmAEDfinal.Text20.Text = Format(vp, "0.00")
'frmAEDfinal.Text21.Text = Format(vdens, "0.00")
'frmAEDfinal.Text22.Text = Format(vsh, "0.00")
'frmAEDfinal.Text23.Text = Format(vtracc, "0.00")
'frmAEDfinal.Text24.Text = Format(velong, "0.00")
'frmAEDfinal.Text25.Text = Format(vmod, "0.00")

frmAEDfinal.Text18.Text = vw
frmAEDfinal.Text19.Text = vd
frmAEDfinal.Text20.Text = vp
frmAEDfinal.Text21.Text = vdens
frmAEDfinal.Text22.Text = vsh
frmAEDfinal.Text23.Text = vtracc
frmAEDfinal.Text24.Text = velong
frmAEDfinal.Text25.Text = vmod

If rs1.Fields("aprobado") = True Then
    frmAEDfinal.Label26.Caption = "Aprobado"
Else
    frmAEDfinal.Label26.Caption = "Desaprobado"
End If



Form1.Enabled = False
Form1.Visible = False
frmAEDfinal.Command2.Enabled = True
frmAEDfinal.Command3.Enabled = True
frmAEDfinal.Show

db.Close
End Sub

Private Sub mnuAEDinic_Click()
Ensayo = InputBox("Ingrese el nº de ensayo que quiere visualizar", "Nº de ensayo")
If Ensayo = "" Then
    Exit Sub
End If

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDorig where ensayo = " & Ensayo)
Set rs1 = db.OpenRecordset("Select * from AEDenv where ensayo = " & Ensayo)

Set rs1 = db.OpenRecordset("Select * from AEDenv where ensayo = " & Ensayo)
If rs.RecordCount = 0 Then
    dsfsdf = MsgBox("No existe el ensayo", vbCritical + vbOKOnly, "Error")
    Exit Sub
Else
    If rs1.RecordCount <> 0 Then
        dsfsdf = MsgBox("Este ensayo ya se encuentra terminado", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
End If



wi = (CDbl(rs.Fields("espesor1")) + CDbl(rs.Fields("espesor2")) + CDbl(rs.Fields("espesor3")) + CDbl(rs.Fields("espesor4")) + CDbl(rs.Fields("espesor5"))) / 5

di = (CDbl(rs.Fields("diamint1")) + CDbl(rs.Fields("diamint2")) + CDbl(rs.Fields("diamint3")) + CDbl(rs.Fields("diamint4")) + CDbl(rs.Fields("diamint5"))) / 5

Pesoi = (CDbl(rs.Fields("peso1")) + CDbl(rs.Fields("peso2")) + CDbl(rs.Fields("peso3")) + CDbl(rs.Fields("peso4")) + CDbl(rs.Fields("peso5"))) / 5

densi = (CDbl(rs.Fields("densidad1")) + CDbl(rs.Fields("densidad2")) + CDbl(rs.Fields("densidad3")) + CDbl(rs.Fields("densidad4")) + CDbl(rs.Fields("densidad5"))) / 5

shi = (CDbl(rs.Fields("dureza1")) + CDbl(rs.Fields("dureza2")) + CDbl(rs.Fields("dureza3")) + CDbl(rs.Fields("dureza4")) + CDbl(rs.Fields("dureza5"))) / 5

tracci = rs.Fields("MPA1")

elongi = rs.Fields("elong1")

modi = rs.Fields("modulo501")

frmAEDfinal.Text1.Text = Ensayo
frmAEDfinal.Text2.Text = Format(wi, "0.00")
frmAEDfinal.Text3.Text = Format(wf, "0.00")
frmAEDfinal.Text4.Text = Format(di, "0.00")
frmAEDfinal.Text5.Text = Format(df, "0.00")
frmAEDfinal.Text6.Text = Format(Pesoi, "0.00")
frmAEDfinal.Text7.Text = Format(pesof, "0.00")
frmAEDfinal.Text8.Text = Format(densi, "0.00")
frmAEDfinal.Text9.Text = Format(densf, "0.00")
frmAEDfinal.Text10.Text = Format(shi, "0.00")
frmAEDfinal.Text11.Text = Format(shf, "0.00")
frmAEDfinal.Text12.Text = Format(tracci, "0.00")
frmAEDfinal.Text13.Text = Format(traccf, "0.00")
frmAEDfinal.Text14.Text = Format(elongi, "0.00")
frmAEDfinal.Text15.Text = Format(elongf, "0.00")
frmAEDfinal.Text16.Text = Format(modi, "0.00")
frmAEDfinal.Text17.Text = Format(modf, "0.00")

frmAEDfinal.Text3.Text = 0
frmAEDfinal.Text5.Text = 0
frmAEDfinal.Text7.Text = 0
frmAEDfinal.Text9.Text = 0
frmAEDfinal.Text11.Text = 0
frmAEDfinal.Text13.Text = 0
frmAEDfinal.Text15.Text = 0
frmAEDfinal.Text17.Text = 0
frmAEDfinal.Text18.Text = 0
frmAEDfinal.Text19.Text = 0
frmAEDfinal.Text20.Text = 0
frmAEDfinal.Text21.Text = 0
frmAEDfinal.Text22.Text = 0
frmAEDfinal.Text23.Text = 0
frmAEDfinal.Text24.Text = 0
frmAEDfinal.Text25.Text = 0



Form1.Enabled = False
Form1.Visible = False
frmAEDfinal.Command2.Enabled = False
frmAEDfinal.Command3.Enabled = False
frmAEDfinal.Show
db.Close
End Sub

Private Sub mnuAEDnuevo_Click()
frmPassword.Show (1)
contr = frmPassword.Password
'contr = InputBox("Ingrese contraseña", "Contraseña")
If contr = "" Then
    Exit Sub
End If

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select * from AEDcontra where contraseña = '" & contr & "'")


frmAEDnuevo.responsablecon = rs.Fields("nombre")





Set rs = db.OpenRecordset("Select * from AEDconst where cod = 0")
Set rs1 = db.OpenRecordset("Select cod from aedconst order by cod desc")

Codigo = rs1.Fields("cod") + 1
frmAEDnuevo.Label3.Caption = Codigo

'frmAEDnuevo.Label19.Caption = ""
'frmAEDnuevo.Label20.Caption = ""
'frmAEDnuevo.Label22.Caption = ""
'frmAEDnuevo.Label23.Caption = ""
'frmAEDnuevo.Label24.Caption = ""
'frmAEDnuevo.Label25.Caption = ""
'frmAEDnuevo.Label26.Caption = ""
'frmAEDnuevo.Label27.Caption = ""

'frmAEDnuevo.Label28.Caption = ""
'frmAEDnuevo.Label29.Caption = ""
'frmAEDnuevo.Label30.Caption = ""
'frmAEDnuevo.Label31.Caption = ""
'frmAEDnuevo.Label32.Caption = ""

'frmAEDnuevo.Label33.Caption = ""
'frmAEDnuevo.Label34.Caption = ""
'frmAEDnuevo.Label35.Caption = ""
'frmAEDnuevo.Label36.Caption = ""


frmAEDnuevo.Label19.Caption = rs.Fields("Titulo")
frmAEDnuevo.Label20.Caption = rs.Fields("Proyecto")
frmAEDnuevo.Label22.Caption = rs.Fields("Doc")
frmAEDnuevo.Label23.Caption = rs.Fields("Especificacion")
frmAEDnuevo.Label24.Caption = rs.Fields("Norma")
frmAEDnuevo.Label25.Caption = rs.Fields("Material")
frmAEDnuevo.Label26.Caption = rs.Fields("Codigo")
frmAEDnuevo.Label27.Caption = rs.Fields("Lote")

frmAEDnuevo.Label28.Caption = rs.Fields("Temperatura")
frmAEDnuevo.Label29.Caption = rs.Fields("Presion")
frmAEDnuevo.Label30.Caption = rs.Fields("Medio")
frmAEDnuevo.Label31.Caption = rs.Fields("Instrumento")
frmAEDnuevo.Label32.Caption = rs.Fields("loteinst")

frmAEDnuevo.Label33.Caption = rs.Fields("Alojam")
frmAEDnuevo.Label34.Caption = 0
frmAEDnuevo.Label35.Caption = rs.Fields("dimen")
frmAEDnuevo.Label36.Caption = rs.Fields("volumen")

frmAEDnuevo.Text1.Text = ""
frmAEDnuevo.Text2.Text = ""
frmAEDnuevo.Text3.Text = ""
frmAEDnuevo.Text4.Text = ""
frmAEDnuevo.Text5.Text = ""
frmAEDnuevo.Text6.Text = ""
frmAEDnuevo.Text7.Text = ""
frmAEDnuevo.Text8.Text = ""
frmAEDnuevo.Text9.Text = ""
frmAEDnuevo.Text10.Text = ""
frmAEDnuevo.Text11.Text = ""
frmAEDnuevo.Text12.Text = ""
frmAEDnuevo.Text13.Text = ""
frmAEDnuevo.Text14.Text = ""
frmAEDnuevo.Text15.Text = ""
frmAEDnuevo.Text16.Text = ""
frmAEDnuevo.Text17.Text = ""
frmAEDnuevo.Text18.Text = ""
frmAEDnuevo.Text19.Text = ""
frmAEDnuevo.Text20.Text = ""
frmAEDnuevo.Text21.Text = ""
frmAEDnuevo.Text22.Text = ""
frmAEDnuevo.Text23.Text = ""
frmAEDnuevo.Text24.Text = ""
frmAEDnuevo.Text25.Text = ""
frmAEDnuevo.Text26.Text = ""
frmAEDnuevo.Text27.Text = ""
frmAEDnuevo.Text28.Text = ""


Form1.Enabled = False
Form1.Visible = False
frmAEDnuevo.Show
db.Close
End Sub

Private Sub mnuAEDseg_Click()
Dim appp As New Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim r As Excel.Range

Set wb = appp.Workbooks.Open("\\Servidor2\e\EntornoBafir\Planillas\SeguimientoAED.xls", , True)
Set ws = wb.Worksheets(1)
ws.PrintOut
DoEvents
sdffsf = MsgBox("Presiones Ok cuando la impresión esté finalizada", vbInformation + vbOKOnly, "Imprimiendo")
wb.Close (False)
End Sub

Private Sub mnuAgregarMedicion_Click()
Me.Enabled = False
frmAgregarMedicion.Text1.Text = ""
frmAgregarMedicion.Show
End Sub

Private Sub mnuAltaConcepto_Click()
Form1.Enabled = False
frmLog.Show (1)
If flags = False Then
    Form1.Enabled = True
    Exit Sub
Else
    a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
    rcset.Open "SELECT clave, permisos FROM flujo_fondos_usuarios where usuario = '" & logUser & "'", CONN, adOpenStatic, adLockReadOnly
    If rcset.RecordCount = 0 Then
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
    If rcset.Fields("clave") = logPass Then
        If rcset.Fields("permisos") = 0 Then
            frmAltaConcepto.Show
            rcset.Close
            Exit Sub
        Else
            asd = MsgBox("Usted no posee permisos para realizar esta operación", vbCritical + vbOKOnly, "Acceso restringido")
            Form1.Enabled = True
            rcset.Close
            Exit Sub
        End If
    Else
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
End If
End Sub

Private Sub mnuAltafondo_Click()
Form1.Enabled = False
frmLog.Show (1)
If flags = False Then
    Form1.Enabled = True
    Exit Sub
Else
    a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
    rcset.Open "SELECT clave, permisos FROM flujo_fondos_usuarios where usuario = '" & logUser & "'", CONN, adOpenStatic, adLockReadOnly
    If rcset.RecordCount = 0 Then
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
    If rcset.Fields("clave") = logPass Then
        If rcset.Fields("permisos") = 0 Then
            rcset.Close
            frmAltaFondos.Show
            frmAltaFondos.Text1.SetFocus
            Exit Sub
        Else
            asd = MsgBox("Usted no posee permisos para realizar esta operación", vbCritical + vbOKOnly, "Acceso restringido")
            Form1.Enabled = True
            rcset.Close
            Exit Sub
        End If
    Else
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
End If
End Sub

Private Sub mnuAM_Click()
    frmShoreAM.Text1.Text = ""
    frmShoreAM.MSFlexGrid1.Clear
    
    frmShoreAM.Show
    Me.Enabled = False
End Sub

Private Sub mnuasientopago_Click()
Call confreg
If confr = "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador decimal y la 'coma' como separador de miles", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Form1.Enabled = False
frmLog.Show (1)
If flags = False Then
    Form1.Enabled = True
    Exit Sub
Else
    a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
    rcset.Open "SELECT clave, permisos FROM flujo_fondos_usuarios where usuario = '" & logUser & "'", CONN, adOpenStatic, adLockReadOnly
    If rcset.RecordCount = 0 Then
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
    If rcset.Fields("clave") = logPass Then
            rcset.Close
            frmAsientoPago.Text1.Text = ""
            frmAsientoPago.Combo1.Clear
            frmAsientoPago.Combo2.Clear
            frmAsientoPago.Text2.Text = ""
            a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
            rcset.Open "SELECT concepto FROM flujo_fondos_concepto order by concepto asc", CONN, adOpenStatic, adLockReadOnly
            Do Until rcset.EOF = True
                frmAsientoPago.Combo1.AddItem (rcset.Fields("concepto"))
                rcset.MoveNext
            Loop
            rcset.Close
            rcset.Open "SELECT fondo FROM flujo_fondos_fondos order by fondo asc", CONN, adOpenStatic, adLockReadOnly
            Do Until rcset.EOF = True
                frmAsientoPago.Combo2.AddItem (rcset.Fields("fondo"))
                rcset.MoveNext
            Loop
            frmAsientoPago.Show
            frmAsientoPago.Text1.SetFocus
    Else
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
End If
End Sub

Private Sub mnuBack_Click()
Me.Enabled = False
frmbackup.Show

'''''''''''''''''''''''''''''''''Todo esto es el programa viejo''''''''''''''''''''''''''''''''''''
'Dim db As Database
'Dim rs As Recordset

'Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select Listados, propietario from backlistados")
'frmBackSel.List1.Clear
'Do Until rs.EOF = True
'frmBackSel.List1.AddItem (rs.Fields("Listados"))
'rs.MoveNext
'Loop
'frmBackSel.Show (1)
'Form1.Enabled = False
'''''''''''''''''
'Set rs = db.OpenRecordset("Select ruta, size from " & frmBackSel.tabla)
'frmBack.List1.Clear
'Dim taman As Double
'If rs.RecordCount = 0 Then
'    frmBack.List1.AddItem ("Lista Vacía")
'Else
'    Do
'        taman = taman + CDbl(rs.Fields("size"))
'        frmBack.List1.AddItem (rs.Fields("ruta"))
'        rs.MoveNext
'    Loop While rs.EOF = False
'End If
'frmBack.Label2.Caption = rs.RecordCount
'db.Close
'frmBack.Height = 3540
'frmBack.Show
'frmBack.Label4.Caption = taman
'Form1.Visible = False
'frmBack.Label6.Caption = frmBackSel.tabla
'''''''''''''''''''''''''''''''''Todo esto es el programa viejo''''''''''''''''''''''''''''''''''''
End Sub

Private Sub mnuBlend_Click()
frmAEDBLEND.mOdO = 1 ' ver en modos de frmaedblend
''''consulta''''
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    
    rst.Open "SELECT batch FROM mezclas where aprobado = TRUE order by batch", cnn, adOpenStatic, adLockReadOnly
''''consulta''''
frmAEDBLEND.List1.Clear
frmAEDBLEND.List2.Clear
frmAEDBLEND.Command4.Enabled = False
frmAEDBLEND.Command5.Enabled = False
frmAEDBLEND.Text1.Text = ""
frmAEDBLEND.Text2.Text = ""
frmAEDBLEND.Check1.Enabled = False
frmAEDBLEND.Check1.Value = 0
Do Until rst.EOF = True
    frmAEDBLEND.List1.AddItem (rst.Fields("batch"))
    rst.MoveNext
Loop



Me.Enabled = False
frmAEDBLEND.Show
rst.Close
cnn.Close

End Sub

Private Sub mnuBusca_Click()
frmBusca.Show


frmBusca.Text9.Clear
frmBusca.Text12.Clear
Dim db As Database
Dim rs2 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs2 = db.OpenRecordset("Select respons_solici from compuestos_para_cotizacion group by respons_solici")
rs2.MoveFirst
Do While rs2.EOF = False
frmBusca.Text9.AddItem (rs2.Fields("respons_solici"))
rs2.MoveNext
Loop
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs2 = db.OpenRecordset("Select cliente from compuestos_para_cotizacion group by cliente order by cliente")
rs2.MoveFirst
Do While rs2.EOF = False
frmBusca.Text12.AddItem (rs2.Fields("cliente"))
rs2.MoveNext
Loop
Form1.Enabled = False
Form1.Visible = False
frmBusca.Text5.Text = ""
frmBusca.Text8.Text = ""
frmBusca.Text9.Text = ""
frmBusca.Text10.Text = ""
frmBusca.Text11.Text = ""
frmBusca.Text12.Text = ""
frmBusca.Text7.Text = ""
frmBusca.Text1.Text = ""
frmBusca.Text2.Text = ""
frmBusca.Text3.Text = ""
frmBusca.Text13.Text = ""
frmBusca.Text4.Text = ""
frmBusca.Text6.Text = ""
frmBusca.Command7.Enabled = False
frmBusca.Command8.Enabled = False
frmBusca.Command2.Enabled = False
frmBusca.Command3.Enabled = True
frmBusca.Command4.Enabled = False
frmBusca.Command5.Enabled = False
End Sub

Private Sub mnuBuscaPart_Click()
frmViscBusca.Show
Form1.Enabled = False
Form1.Visible = False
End Sub

Private Sub mnubuscarcuerda_Click()
frmcuerdasbusca.Combo1.Clear
frmcuerdasbusca.Combo2.Clear
Form1.Enabled = False
Form1.Visible = False

Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select proveedor from cuerdas group by proveedor")
Set rs1 = db.OpenRecordset("Select Cuerda from cuerdas group by cuerda")
Do Until rs.EOF = True
frmcuerdasbusca.Combo2.AddItem (rs.Fields("proveedor"))
rs.MoveNext
Loop
Do Until rs1.EOF = True
frmcuerdasbusca.Combo1.AddItem (rs1.Fields("cuerda"))
rs1.MoveNext
Loop
db.Close
frmcuerdasbusca.Show
End Sub

Private Sub mnuBuscaValores_Click()
frmBuscaTraccion.Show
frmBuscaTraccion.Height = 5565
frmBuscaTraccion.Command5.Visible = False
frmBuscaTraccion.Command2.Visible = True
Form1.Enabled = False
Form1.Visible = False
End Sub

Private Sub mnuCargarLote_Click()
frmCargarLote.buscar = False
'frmCargarLote.Label2.Caption = Format(Date, "DD/MM/YY")
'frmCargarLote.ReservaLote
frmCargarLote.Show
Me.Enabled = False

End Sub

Private Sub mnucargarlotebuscar_Click()
frmcargarlotebusqueda.Label2.Caption = ""
frmcargarlotebusqueda.Label4.Caption = ""
frmcargarlotebusqueda.Text1.Text = ""
frmcargarlotebusqueda.Text2.Text = ""
frmcargarlotebusqueda.Text3.Text = ""
frmcargarlotebusqueda.Text4.Text = ""
frmcargarlotebusqueda.Text5.Text = ""
frmcargarlotebusqueda.Label13.Caption = ""
frmcargarlotebusqueda.Text6.Text = ""
frmcargarlotebusqueda.Label14.Caption = ""
frmcargarlotebusqueda.Text7.Text = ""



frmcargarlotebusqueda.Show
Me.Enabled = False
End Sub

Private Sub mnuCertificado_Click()
ShellExecute 0&, vbNullString, "\\Servidor2\e\ISO9000\MANUAL DE CALIDAD\Certificado de Calidad ISO 9001-2000.pdf", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub mnuCompe_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Dim Codigo As String
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Codigo = InputBox("Ingrese el codigo de ensayo", "Codigo")
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select codigo_ensayo, original1, original2, espaciador from compresion_temporal where codigo_ensayo = '" & Codigo & "'")
Set rs1 = db.OpenRecordset("Select codigo_ensayo, compuesto, partida, tiempo_temperatura, compresion_porc from compresion where codigo_ensayo = " & Codigo & "")
If rs.RecordCount = 0 Then
sadghsfh = MsgBox("No existe el ensayo indicado", vbCritical + vbOKOnly, "Error")
db.Close
Exit Sub
End If
env1 = InputBox("Ingrese el valor envejecido 1/2", "Valor envejecido 1/2 " & rs1.Fields("codigo_ensayo") & " " & rs1.Fields("Compuesto") & " " & rs1.Fields("partida") & " " & rs1.Fields("tiempo_temperatura"))
If env1 = "" Then
    Exit Sub
End If
env2 = InputBox("Ingrese el valor envejecido 2/2", "Valor envejecido 2/2" & rs1.Fields("codigo_ensayo") & " " & rs1.Fields("Compuesto") & " " & rs1.Fields("partida") & " " & rs1.Fields("tiempo_temperatura"))
If env2 = "" Then
    Exit Sub
End If
a = InStr(1, env1, ".")
    If a <> 0 Then
        Mid(env1, a) = ","
    End If
b = InStr(1, env2, ".")
    If b <> 0 Then
        Mid(env2, b) = ","
    End If
    promediooriginales = CDbl((rs.Fields("original1")) + CDbl(rs.Fields("original2"))) / 2
    promedioenvejecidos = (CDbl(env1) + CDbl(env2)) / 2
    dac = (promediooriginales - promedioenvejecidos)
    dbc = (promediooriginales - CDbl(rs.Fields("espaciador")))
    compresion = Format(dac / dbc * 100, "0.00")
    rs1.Edit
    rs1.Fields("compresion_porc") = compresion
    rs1.Update
    db.Close
End Sub

Private Sub mnuCompo_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim comPuesto As String
Dim parTida As String
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
cantidadcomp = InputBox("Ingrese cantidad de compresiones a ingresar", "Compresiones")
cantidadcomp = CInt(cantidadcomp)
For i = 1 To cantidadcomp
Set rs = db.OpenRecordset("Select codigo_ensayo, original1, original2, espaciador from compresion_temporal")
Set rs1 = db.OpenRecordset("Select fecha, codref,compresion, probeta, codigo_ensayo, compuesto, partida, tiempo_temperatura, compresion_porc from compresion order by codigo_ensayo")
comPuesto = InputBox("Ingrese compuesto", "Compuesto")
If comPuesto = "" Then
    Exit Sub
End If
parTida = InputBox("Ingrese Partida", "Partida")
If parTida = "" Then
    Exit Sub
End If
rs1.MoveLast
Codigo = CInt(rs1.Fields("codigo_ensayo")) + 1

frmSelProbeta.Show (1)
probeta = frmSelProbeta.probeta

Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset



sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT referencia FROM ensayos where tipo = 'compresion' order by referencia ", cnn, adOpenStatic, adLockReadOnly
    frmSeleccionarEnsayo.List1.Clear
    If Not rst.RecordCount = 0 Then
    Do Until rst.EOF = True
        frmSeleccionarEnsayo.List1.AddItem (rst.Fields("Referencia"))
        rst.MoveNext
    Loop
    
    End If
    rst.Close
    frmSeleccionarEnsayo.Show (1)

    tiempoytemperatura = frmSeleccionarEnsayo.Ensayo
'tiempoYtemperatura = InputBox("Ingrese el tiempo y la temperatura del ensayo", "Tiempo y Temperatura")
'If tiempoYtemperatura = "" Then
'    Exit Sub
'End If

valororiginal1 = InputBox("Ingrese el espesor original 1", "Compresion - Valor Original 1/2")
If valororiginal1 = "" Then
    Exit Sub
End If
valororiginal2 = InputBox("Ingrese el espesor original 2", "Compresion - Valor Original 2/2")
If valororiginal2 = "" Then
    Exit Sub
End If
    a = InStr(1, valororiginal1, ".")
    If a <> 0 Then
        Mid(valororiginal1, a) = ","
    End If
    b = InStr(1, valororiginal2, ".")
    If b <> 0 Then
        Mid(valororiginal2, b) = ","
    End If

espaciador = InputBox("Ingrese el valor del espaciador (mm)", "Espaciador")
Do
    durezashorea = InputBox("Ingrese la dureza Shore A", "Shore A")
Loop Until IsNumeric(durezashorea)

Do
    durezashorem = InputBox("Ingrese la dureza Shore M", "Shore M")
Loop Until IsNumeric(durezashorem)


If espaciador = "" Then
    Exit Sub
End If
j = InStr(1, espaciador, ".")
    If j <> 0 Then
        Mid(espaciador, j) = ","
    End If
origcomp = (CDbl(valororiginal1) + CDbl(valororiginal2)) / 2
j = InStr(1, origcomp, ".")
    If j <> 0 Then
        Mid(origcomp, j) = ","
    End If
compresion = Format(((espaciador - origcomp) / -origcomp) * 100, "#.##") & "%"



rst.Open "SELECT codigo FROM ensayos where referencia = '" & frmSeleccionarEnsayo.Ensayo & "' and tipo = 'Compresion'", cnn, adOpenStatic, adLockReadOnly
codref = rst.Fields("codigo")

rs.AddNew
rs1.AddNew
rs.Fields("codigo_ensayo") = Codigo
rs1.Fields("codigo_ensayo") = Codigo
rs1.Fields("probeta") = probeta
rs1.Fields("compresion") = compresion
rs1.Fields("fecha") = Date
rs1.Fields("compuesto") = comPuesto
rs1.Fields("partida") = parTida
rs1.Fields("codref") = codref
rs1.Fields("Tiempo_temperatura") = tiempoytemperatura
rs.Fields("original1") = valororiginal1
rs.Fields("original2") = valororiginal2
rs.Fields("espaciador") = espaciador
rs1.Fields("compresion_porc") = "0"
rs.Update
rs1.Update
rst.Close
If probeta = "ASTM D395 1A (12,5x29) SIN APILAR" Then
    rst.Open "SELECT * FROM espesores", cnn, adOpenStatic, adLockOptimistic
    rst.AddNew
    rst.Fields("fecha") = Date
    rst.Fields("compuesto") = comPuesto
    rst.Fields("partida") = parTida
    rst.Fields("espesor") = (CDbl(valororiginal1) + CDbl(valororiginal2)) / 2
    rst.Update
    rst.Close
End If
    


rst1.Open "SELECT * FROM shoream", cnn, adOpenStatic, adLockOptimistic
rst1.AddNew
rst1.Fields("shorea") = durezashorea
rst1.Fields("shorem") = durezashorem
rst1.Fields("fecha") = Date
rst1.Fields("compuesto") = comPuesto
rst1.Fields("partida") = parTida
rst1.Update
rst1.Close

dasd = MsgBox("El codigo perteneciente al compuesto " & comPuesto & "-" & parTida & " es " & Codigo, vbInformation + vbOKOnly, "Codigo de compresión")
Next
db.Close
cnn.Close
End Sub

Private Sub mnuconsultbuscar_Click()
Form1.Enabled = False
Form1.Visible = False
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select cliente from Consultas group by cliente")

Do Until rs.EOF = True
frmconsultabusca.Combo1.AddItem (rs.Fields("cliente"))
rs.MoveNext
Loop

db.Close



frmconsultabusca.Show
End Sub

Private Sub mnuconsulting_Click()

frmConsultas.Text1.Text = ""
frmConsultas.Text2.Text = ""
frmConsultas.Text3.Text = ""
frmConsultas.Combo1.Text = ""
frmConsultas.Combo1.Clear
frmConsultas.Combo2.Text = ""
frmConsultas.Combo2.Clear
frmConsultas.Text4.Text = ""
frmConsultas.Text5.Text = ""
frmConsultas.Combo3.Clear
frmConsultas.Combo4.Clear
frmConsultas.Combo3.Text = ""
frmConsultas.Combo4.Text = ""
frmConsultas.Check1.Value = False
frmConsultas.Check2.Value = False
frmConsultas.Check3.Value = False
frmConsultas.Check4.Value = False

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT codigo from consultas order by codigo desc")

If rs.RecordCount = 0 Then
    codigonuevo = 1
Else
    codigonuevo = rs.Fields("codigo") + 1
End If

frmConsultas.Text1.Text = codigonuevo
frmConsultas.Text2.Text = Date

Set rs = db.OpenRecordset("SELECT cliente from consultas group by cliente")
If rs.RecordCount <> 0 Then
    Do Until rs.EOF = True
        frmConsultas.Combo1.AddItem (rs.Fields("cliente"))
        rs.MoveNext
    Loop
End If

Set rs = db.OpenRecordset("SELECT responsable_cons from consultas group by responsable_cons")
If rs.RecordCount <> 0 Then
    Do Until rs.EOF = True
        frmConsultas.Combo2.AddItem (rs.Fields("responsable_cons"))
        rs.MoveNext
    Loop
End If


db.Close
Form1.Enabled = False
Form1.Visible = False
frmConsultas.Show
frmConsultas.Combo1.SetFocus


End Sub

Private Sub mnuconsultmodcons_Click()
Form1.Enabled = False
Form1.Visible = False
frmConsultaModCons.Text1.Text = ""
frmConsultaModCons.Text2.Text = ""
frmConsultaModCons.Text3.Text = ""
frmConsultaModCons.Text4.Text = ""
frmConsultaModCons.Text5.Text = ""
frmConsultaModCons.Text6.Text = ""
frmConsultaModCons.Show
End Sub

Private Sub mnuconsultrespond_Click()
frmConsultaResp.Text1.Text = ""
frmConsultaResp.Text2.Text = ""
frmConsultaResp.Text3.Text = ""
frmConsultaResp.Text4.Text = ""
frmConsultaResp.Text5.Text = ""
frmConsultaResp.Text6.Text = ""
frmConsultaResp.Text7.Text = ""
frmConsultaResp.Combo3.Text = ""
frmConsultaResp.Text9.Text = ""
frmConsultaResp.Text10.Text = ""
frmConsultaResp.Text11.Text = ""
frmConsultaResp.Combo1.Text = ""
frmConsultaResp.Combo2.Text = ""

Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT responsable_resp FROM consultas group by responsable_resp")

If rs.RecordCount <> 0 Then
    Do Until rs.EOF = True
        frmConsultaResp.Combo3.AddItem (rs.Fields("responsable_resp") & "")
        rs.MoveNext
    Loop
End If

Form1.Enabled = False
Form1.Visible = False
frmConsultaResp.Show
db.Close
End Sub

Private Sub mnuConsumos_Click()
Me.Enabled = False
frmConsumos.Show
End Sub

Private Sub mnuContra_Click()
Dim db As Database
Dim rs As Recordset
frmPassword.Show (1)
conu = frmPassword.Password
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM REG where funcion = 'godkey'")
If conu <> rs.Fields("dato") Then
    asdsdf = MsgBox("Contraseña incorrecta", vbCritical + vbOKOnly, "Error")
    Exit Sub
Else
frmLogger.Show
End If
End Sub

Private Sub mnuContracc_Click()
   frmespesores.Combo1.Clear
   Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset

    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
   
   sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

With cnn
'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
.Open
End With

rst.Open "SELECT N_FORMULA FROM Formbase", cnn, adOpenStatic, adLockReadOnly

        rst.MoveFirst
        Do Until rst.EOF = True
            frmespesores.Combo1.AddItem (rst.Fields("N_FORMULA"))
            rst.MoveNext
        Loop

   
   
   frmespesores.Show
   Me.Enabled = False
End Sub

Private Sub mnucostoscomp_Click()
Call confreg
If confr = "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador decimal y la 'coma' como separador de miles", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Me.Enabled = False
frmSimulacionCostos.Show
End Sub

Private Sub mnuCrearReg_Click()

Dim db1 As Database
Dim rs1 As Recordset
Dim td1 As TableDef
Dim fld1 As Field
Dim strFecha As String
Dim db As Database
Dim rs As Recordset
Dim intFields As Integer
Dim intRs As Integer
Dim strForm As String
Dim strPass As String
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT N_FORMULA, COSTO_TOTAL FROM Formbase")
Set db1 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set td1 = db1.TableDefs("HISTORICO_PRECIOS")
Set rs1 = db1.OpenRecordset("REG", dbOpenTable)
rs1.Index = "primarykey"
rs1.Seek "=", "HistKey"
strPass = rs1.Fields("dato")
'''''''''''''''esta era la validacion antes sacar cuando se active login
ingr = InputBox("Ingrese la clave", "Clave de ingreso")
If ingr <> strPass Then
    roro = MsgBox("La clave no es valida", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
'''''''''''''''esta era la validacion antes sacar cuando se active login
strFecha = CStr(Date)
intFields = (td1.Fields.Count) - 1
For contador = 1 To intFields
    If strFecha = td1.Fields(contador).Name Then
        resp = MsgBox("No se puede crear el registro por que ya está creado", vbOKOnly + vbCritical, "Error")
        db.Close
        db1.Close
        Exit Sub
    End If
Next
Set fld1 = td1.CreateField(strFecha, dbText)
td1.Fields.Append fld1
Set rs1 = db1.OpenRecordset("HISTORICO_PRECIOS", dbOpenTable)
rs.MoveFirst
rs1.MoveFirst
Do While rs.EOF <> True
    Y = rs.RecordCount
    On Error Resume Next
    strForm = rs.Fields("N_formula")
   Erri = Err.Number
   If Erri = 94 Then
   Exit Do
   End If
    rs1.Index = "primarykey"
    rs1.Seek "=", strForm
    If rs1.NoMatch Then
        With rs1
        .AddNew
        .Fields("N_FORMULA") = strForm
        .Fields(strFecha) = rs.Fields("COSTO_TOTAL").Value
        
        .Update
        End With
        col = rs1.Fields.Count
        For col1 = 1 To (col - 2)
        rs1.Seek "=", strForm
        rs1.Edit
        rs1.Fields(col1) = 0
        rs1.Update
        Next
    Else
    rs1.Edit
    rs1.Fields(strFecha) = rs.Fields("COSTO_TOTAL").Value
    rs1.Update
    End If
    rs.MoveNext
Loop
db.Close
reso = MsgBox("Se ha grabado el registro para la fecha " & strFecha, vbInformation)
Set rs1 = db1.OpenRecordset("REG", dbOpenTable)
rs1.Index = "primarykey"
rs1.Seek "=", "HistFecha"
rs1.Edit
rs1.Fields("Dato") = strFecha
rs1.Update
db1.Close
ReDim destinatarios(1 To 3)
indicedestinatarios = 3
asunto = "Entorno Bafir: Actualización de Historico de precios"
mail = "Se ha cargado el registro correspondiente a la fecha " & Date & " del histórico de precios"
destinatarios(1) = "pablopirri@bafir.com.ar"
destinatarios(2) = "laboratorio@bafir.com.ar"
'destinatarios(3) = "1140706885@sms.movistar.net.ar"
destinatarios(3) = "entornobafir@gmail.com"
frmSendinfo.Show
frmSendinfo.Hide
End Sub

Private Sub mnuCuerdaInforme_Click()
frmCuerdaInforme.Combo1.Clear
frmCuerdaInforme.Combo2.Clear
frmCuerdaInforme.Combo3.Clear
frmCuerdaInforme.Check1.Value = False
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select proveedor from cuerdas group by proveedor")
rs.MoveFirst
Do Until rs.EOF = True
frmCuerdaInforme.Combo1.AddItem (rs.Fields("proveedor"))
rs.MoveNext
Loop
db.Close
Form1.Enabled = False
Form1.Visible = False
frmCuerdaInforme.Show
frmCuerdaInforme.Check1.Value = 0
frmCuerdaInforme.Label4.Caption = ""
frmCuerdaInforme.Combo1.SetFocus
End Sub

Private Sub mnuDatos_Click()
Command3.Visible = False
mnuDatos.Checked = True
mnuHist.Checked = False
mnuCrearReg.Enabled = True
Text7.Visible = True
Label10.Visible = True
Label13.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text8.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label9.Visible = True
MSFlexGrid1.Visible = False
MSFlexGrid1.Clear
mnut90.Enabled = True
End Sub
Private Sub mnuDens_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
frmDens.Show
frmDens.Enabled = True
Form1.Enabled = False
Form1.Visible = False
End Sub

Private Sub mnuDesgarroBuscar_Click()
frmDesgarroBuscar.Text1.Text = ""
frmDesgarroBuscar.MSFlexGrid1.Clear
Form1.Enabled = False
Form1.Visible = False
frmDesgarroBuscar.Show
frmDesgarroBuscar.Text1.SetFocus
End Sub

Private Sub mnuDesgarrosDimensiones_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Form1.Enabled = False
Form1.Visible = False
frmSelProbeta.Show (1)
probeta = frmSelProbeta.probeta
compuestos = InputBox("Ingrese Nombre de compuesto", "Compuesto")
parTIIda = InputBox("Ingrese partida", "Partida")
cantidad = InputBox("Ingrese cantidad de probetas a ensayar", "Probetas")
If cantidad = "" Then
    Exit Sub
End If
If Not IsNumeric(cantidad) Then
    sdfsdf = MsgBox("Debe ingresar un número válido", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
cantidad = CInt(cantidad)
Dim espesores As String
For agregar = 1 To cantidad
    ingreso = punto_por_coma(InputBox("Ingrese espesor " & agregar & "/" & cantidad, "Ingreso de espesores"))
    
    If agregar = 1 Then
    espesores = ingreso & "@"
    End If
    If agregar <> 1 And agregar <> cantidad Then
    espesores = espesores & ingreso & "@"
    End If

    If agregar = cantidad Then
    espesores = espesores & ingreso
    End If
Next
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM desgarros order by ensayo")
Set rs1 = db.OpenRecordset("SELECT * FROM desgarros_temp order by ensayo")

If rs.RecordCount = 0 Then
    ensayOO = 1
Else
    rs.MoveLast
    ensayOO = rs.Fields("ensayo") + 1
End If

rs.AddNew
rs.Fields("ensayo") = ensayOO
rs.Fields("compuesto") = compuestos
rs.Fields("partida") = parTIIda
rs.Fields("probeta") = probeta
rs.Fields("fecha") = Date
rs1.AddNew
rs1.Fields("ensayo") = ensayOO
rs1.Fields("espesores") = espesores
rs.Update
rs1.Update
db.Close
dfsdfs = MsgBox("Valores guardados como ensayo " & ensayOO, vbInformation + vbOKOnly, "Desgarros")
Form1.Enabled = True
Form1.Visible = True
End Sub

Private Sub mnuDesgarrosTraccion_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Ensayo = InputBox("Ingrese código de ensayo", "Desgarro")
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM desgarros where ensayo = " & Ensayo)
Set rs1 = db.OpenRecordset("SELECT * FROM desgarros_temp where ensayo = " & Ensayo)
If rs.RecordCount = 0 Then
    sdfsdf = MsgBox("No existe el ensayo indicado", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If rs1.RecordCount = 0 Then
    sdfsdf = MsgBox("El ensayo se encuentra cerrado", vbCritical + vbOKOnly, "Ensayo cerrado")
    Exit Sub
End If

espesores = Explode("@", rs1.Fields("espesores"))
db.Close
probet = UBound(espesores)

ReDim kiLos(probet)
For i = 0 To probet
    kiLos(i) = punto_por_coma(InputBox("Ingrese tracción " & i + 1 & "/" & probet + 1, "Desgarro " & i + 1 & "/" & probet + 1))
Next
''''''''''''''''''desgarro = 0
ReDim desgarro(probet)
For i = 0 To probet
''''''''''''''''''desgarro = desgarro + (CDbl(kiLos(i)) / CDbl(espesores(i)))
    desgarro(i) = CDbl(kiLos(i)) / CDbl(espesores(i))
Next
''''''''''''''''''desgarro = desgarro / (probet + 1)

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT * FROM desgarros where ensayo = " & Ensayo)
Set rs1 = db.OpenRecordset("SELECT * FROM desgarros_temp where ensayo = " & Ensayo)
rs.Edit
'''''''''''''''''rs.Fields("Valor") = Format(desgarro, "0.000") & " kg/mm"
Valor = Implode(desgarro)

rs.Fields("Valor") = Valor
rs.Update

rs1.Delete

sdfsdfsf = MsgBox("Se ha guardado el valor para el ensayo " & Ensayo & ". Se ha cerrado el ensayo. Los valores de desgarro se almacenan en Kg/mm. Considerelo para la realización de los calculos", vbInformation + vbOKOnly, "Cerrando ensayo")
db.Close
End Sub

Private Sub mnuDespachar_Click()
frmDespachar.Combo1.Clear
frmDespachar.Combo2.Clear
frmDespachar.Text1.Text = ""
frmDespachar.Text2.Text = ""


Dim cnn As ADODB.Connection
Dim cnn1 As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set cnn1 = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\centralpesado.mdb"

    With cnn
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

    rst.Open "SELECT descrip FROM producto order by descrip asc", cnn, adOpenStatic, adLockReadOnly
       
    Do Until rst.EOF = True
        frmDespachar.Combo1.AddItem (rst.Fields("descrip"))
        rst.MoveNext
    Loop
    rst.Close
    frmDespachar.Combo2.AddItem ("Causer")
    frmDespachar.Combo2.AddItem ("Campex")
    frmDespachar.Combo2.AddItem ("ISISA")
    Me.Enabled = False
    frmDespachar.Show
    cnn.Close
End Sub

Private Sub mnuDimensiones_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    






Form1.Enabled = False
Form1.Visible = False
frmTraccion.Show
frmTraccion.Check1.Value = 0
frmTraccion.Check1.Caption = ""
frmTraccion.Combo1.Text = ""
frmTraccion.Text1.Text = ""
frmTraccion.Text2.Text = ""
frmTraccion.Label6.Caption = ""
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select COMPUESTO from TRACCION GROUP BY COMPUESTO")
If rs.RecordCount <> 0 Then
rs.MoveFirst
Do Until rs.EOF = True
frmTraccion.Combo1.AddItem (rs.Fields("compuesto"))
rs.MoveNext
Loop
End If


'Set rs = db.OpenRecordset("Select referencia from TRACCION GROUP BY referencia")
rst.Open "SELECT referencia FROM ensayos where tipo = 'Envejecimiento' group by referencia", cnn, adOpenStatic, adLockReadOnly
'rs.MoveFirst
frmTraccion.Text2.Clear
Do While rst.EOF = False
frmTraccion.Text2.AddItem (rst.Fields("referencia"))
rst.MoveNext
Loop
frmTraccion.Text2.Enabled = False
db.Close
cnn.Close
End Sub




Private Sub mnuDolar_Click()
    'ShellExecute frmDolar.Hwnd, "Open", "http://www.midolar.com.ar/dolar.xml", &O0, &O0, SW_NORMAL
    'frmDolar.Show
    backupear_restantes ("Z:\Backups_del_area_industrial\20090219164533_Backuplaboratorio.log")
End Sub
    
Private Sub mnuEnsayoAgregarNuevo_Click()
Me.Enabled = False

Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

    
    rst.Open "SELECT codigo FROM ensayos order by codigo desc", cnn, adOpenStatic, adLockReadOnly
    codigoens = rst.Fields("codigo") + 1
    

frmEnsayoAgregarNuevo.Text1.Text = codigoens
frmEnsayoAgregarNuevo.List1.Clear
frmEnsayoAgregarNuevo.List2.Clear
frmEnsayoAgregarNuevo.List1.AddItem ("Envejecimiento")
frmEnsayoAgregarNuevo.List1.AddItem ("Compresion")
frmEnsayoAgregarNuevo.List1.AddItem ("Varios")
    rst.Close
    rst.Open "SELECT nombre FROM stock where tipo = 'Fluidos' or tipo = 'Reactivos' order by nombre asc", cnn, adOpenStatic, adLockReadOnly

frmEnsayoAgregarNuevo.List2.AddItem "Aire"
Do Until rst.EOF = True
    frmEnsayoAgregarNuevo.List2.AddItem (rst.Fields("Nombre"))
    rst.MoveNext
Loop




cnn.Close
frmEnsayoAgregarNuevo.Command1.Visible = True
frmEnsayoAgregarNuevo.Command2.Visible = True
frmEnsayoAgregarNuevo.Command3.Visible = False
frmEnsayoAgregarNuevo.Command4.Visible = False
frmEnsayoAgregarNuevo.Command5.Visible = False
frmEnsayoAgregarNuevo.Command6.Visible = False
frmEnsayoAgregarNuevo.Show
frmEnsayoAgregarNuevo.Text2.SetFocus
End Sub

Private Sub mnuEnsayosVer_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT * FROM ensayos", cnn, adOpenStatic, adLockReadOnly
    
    frmEnsayosVer.MSFlexGrid1.Cols = 3
    frmEnsayosVer.MSFlexGrid1.Rows = rst.RecordCount + 1
    
    frmEnsayosVer.MSFlexGrid1.TextMatrix(0, 0) = "Código"
    frmEnsayosVer.MSFlexGrid1.TextMatrix(0, 1) = "Referencia"
    frmEnsayosVer.MSFlexGrid1.TextMatrix(0, 2) = "Tipo"
    
    For i = 1 To rst.RecordCount
        frmEnsayosVer.MSFlexGrid1.TextMatrix(i, 0) = rst.Fields("codigo")
        frmEnsayosVer.MSFlexGrid1.TextMatrix(i, 1) = rst.Fields("Referencia")
        frmEnsayosVer.MSFlexGrid1.TextMatrix(i, 2) = rst.Fields("Tipo")
        rst.MoveNext
    Next
    cnn.Close
    AutoGrid frmEnsayosVer.MSFlexGrid1


Me.Enabled = False
frmEnsayosVer.Show
End Sub

Private Sub mnuEnsExt_Click()
Form1.Enabled = False
Form1.Visible = False
frmEnsExt.Label5.Caption = ""
frmEnsExt.Label5.Caption = Date
frmEnsExt.Combo1.Clear
frmEnsExt.Combo2.Clear
frmEnsExt.Combo3.Clear
frmEnsExt.Combo4.Clear
frmEnsExt.Text1.Text = ""
frmEnsExt.Text2.Text = ""
frmEnsExt.MSFlexGrid1.Clear

Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select compuesto from ensayos_externos group by compuesto")

If rs.RecordCount <> 0 Then

    rs.MoveFirst
    Do Until rs.EOF = True
    frmEnsExt.Combo4.AddItem (rs.Fields("compuesto"))
    rs.MoveNext
    Loop

    Set rs = db.OpenRecordset("Select ensayo from ensayos_externos group by ensayo")
    rs.MoveFirst
    Do Until rs.EOF = True
    frmEnsExt.Combo1.AddItem (rs.Fields("ensayo"))
    rs.MoveNext
    Loop

    Set rs = db.OpenRecordset("Select norma_cond from ensayos_externos group by norma_cond")
    rs.MoveFirst
    Do Until rs.EOF = True
    frmEnsExt.Combo2.AddItem (rs.Fields("norma_cond"))
    rs.MoveNext
    Loop

    Set rs = db.OpenRecordset("Select proveedor from ensayos_externos group by proveedor")
    rs.MoveFirst
    Do Until rs.EOF = True
    frmEnsExt.Combo3.AddItem (rs.Fields("proveedor"))
    rs.MoveNext
    Loop
    Else
    frmEnsExt.Command2.Enabled = False
    End If
    db.Close
frmEnsExt.Show
frmEnsExt.Combo4.SetFocus
End Sub

Private Sub mnufd_Click()
    Me.Enabled = False
    verformUlasD
End Sub

Private Sub mnuFluid_Click()
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
If Form1.Caption <> "Entorno Bafir - Entorno de Gestión de planta - Modo Laboratorio" Then
Dim db1 As Database
Dim rs1 As Recordset
Set db1 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db1.OpenRecordset("REG", dbOpenTable)
rs1.Index = "primarykey"
rs1.Seek "=", "LabKey"
strPass = rs1.Fields("dato")
ingr = InputBox("Ingrese la clave", "Clave de ingreso")
If ingr <> strPass Then
    roro = MsgBox("La clave no es valida", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Else
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
Form1.Enabled = False
frmFluidos.Show
Form1.Visible = False
frmFluidos.Height = 2085
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
End If
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
End Sub

Private Sub mnuFormulas_Click()
'Call actualiza_Formulas
End Sub

Private Sub mnuHist_Click()
mnuDatos.Checked = False
mnuHist.Checked = True
Text7.Visible = False
Label10.Visible = False
Label13.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text8.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label9.Visible = False
MSFlexGrid1.Visible = True
mnut90.Enabled = False
mnuCrearReg.Enabled = False
Command3.Visible = True
End Sub

Private Sub mnuIndicadores_Click()
Me.Enabled = False
frmIndicadores.Show
End Sub

Private Sub mnuIngFormula_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
    frmDesarrFormulaIng.Combo1(0).Text = ""
    frmDesarrFormulaIng.Text1(0).Text = ""
    frmDesarrFormulaIng.Combo2(0).Text = ""
    Dim db As Database
    Dim rs As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select cod_prod,descrip from producto order by descrip")
    
    codigoFor = InputBox("Ingrese Codigo de fórmula")
    componentes = InputBox("Ingrese cantidad de componentes")
    
    componentes = componentes - 1
    
    Do Until rs.EOF = True
            frmDesarrFormulaIng.Combo1(0).AddItem (rs.Fields("descrip"))
            rs.MoveNext
    Loop
        rs.MoveFirst
    frmDesarrFormulaIng.Combo1(0).TabIndex = 0
    frmDesarrFormulaIng.Text1(0).TabIndex = 1
    frmDesarrFormulaIng.Combo2(0).TabIndex = 2
    
    
    For i = 1 To componentes
        taBi = taBi + 3
        Load frmDesarrFormulaIng.Text1(i)
        Load frmDesarrFormulaIng.Combo1(i)
        Load frmDesarrFormulaIng.Combo2(i)
        
        
        frmDesarrFormulaIng.Text1(i).Visible = True
        Do Until rs.EOF = True
            frmDesarrFormulaIng.Combo1(i).AddItem (rs.Fields("descrip"))
            rs.MoveNext
        Loop
            rs.MoveFirst
        frmDesarrFormulaIng.Combo1(i).Visible = True
        frmDesarrFormulaIng.Combo2(i).Visible = True
        frmDesarrFormulaIng.Combo2(i).AddItem ("A")
        frmDesarrFormulaIng.Combo2(i).AddItem ("B")
        frmDesarrFormulaIng.Text1(i).Top = frmDesarrFormulaIng.Text1(i - 1).Top + frmDesarrFormulaIng.Text1(i).Height
        frmDesarrFormulaIng.Combo1(i).Top = frmDesarrFormulaIng.Combo1(i - 1).Top + frmDesarrFormulaIng.Combo1(i).Height
        frmDesarrFormulaIng.Combo2(i).Top = frmDesarrFormulaIng.Combo2(i - 1).Top + frmDesarrFormulaIng.Combo1(i).Height
    Next
    Form1.Visible = False
    Form1.Enabled = False
    db.Close
    frmDesarrFormulaIng.Label5.Caption = codigoFor
    frmDesarrFormulaIng.Label12.Caption = ""
    frmDesarrFormulaIng.Label13.Caption = ""
    frmDesarrFormulaIng.Label7.Caption = ""
    taBi = 0
    For i = 1 To componentes
        taBi = taBi + 3
        frmDesarrFormulaIng.Combo1(i).TabIndex = taBi
        frmDesarrFormulaIng.Text1(i).TabIndex = taBi + 1
        frmDesarrFormulaIng.Combo2(i).TabIndex = taBi + 2
    Next
    frmDesarrFormulaIng.Command2.Enabled = True
    frmDesarrFormulaIng.Command4.Enabled = True
    frmDesarrFormulaIng.Command5.Enabled = True
    frmDesarrFormulaIng.Command6.Enabled = True
    frmDesarrFormulaIng.Command3.Enabled = True
    frmDesarrFormulaIng.Command7.Enabled = True
    frmDesarrFormulaIng.Show
End Sub

Private Sub mnuIngNoconfReo_Click()
Form1.Enabled = False
Form1.Visible = False
frmReomNIn.Flag = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select compuesto from noconf_reometro group by compuesto")
Set rs1 = db.OpenRecordset("Select motivo from noconf_reometro group by motivo")
Set rs2 = db.OpenRecordset("Select resolucion from noconf_reometro group by resolucion")
Set rs3 = db.OpenRecordset("Select accion_posterior from noconf_reometro group by accion_posterior")

frmReomNIn.Combo1.Clear
frmReomNIn.Combo2.Clear
frmReomNIn.Combo3.Clear
frmReomNIn.Combo4.Clear

If rs.RecordCount = 0 Then
Else
rs.MoveFirst
Do Until rs.EOF = True
frmReomNIn.Combo1.AddItem (rs.Fields("compuesto"))
rs.MoveNext
Loop
End If

If rs1.RecordCount = 0 Then
Else
rs1.MoveFirst
Do Until rs1.EOF = True
frmReomNIn.Combo2.AddItem (rs1.Fields("motivo"))
rs1.MoveNext
Loop
End If

If rs2.RecordCount = 0 Then
Else
rs2.MoveFirst
Do Until rs2.EOF = True
frmReomNIn.Combo3.AddItem (rs2.Fields("resolucion"))
rs2.MoveNext
Loop
End If

If rs3.RecordCount = 0 Then
Else
rs3.MoveFirst
Do Until rs3.EOF = True

frmReomNIn.Combo4.AddItem (rs3.Fields("accion_posterior") & "")

rs3.MoveNext
Loop
End If
db.Close
frmReomNIn.Text1.Text = Date
frmReomNIn.Text3.Text = ""
frmReomNIn.Text4.Text = ""
frmReomNIn.Text5.Text = ""
frmReomNIn.Text2.Text = ""
frmReomNIn.Combo2.Text = ""
frmReomNIn.Combo3.Text = ""
frmReomNIn.Combo4.Text = ""
frmReomNIn.Check1.Value = 0
frmReomNIn.Check2.Value = 0
frmReomNIn.Command4.Visible = False
frmReomNIn.Command5.Visible = False
formUlario = Me.Name
frmReomNIn.Show

End Sub

Private Sub mnuIngVisc_Click()
Form1.Enabled = False
Form1.Visible = False
frmViscIng.Show
frmViscIng.Visible = False
frmViscIng.Text5.Text = ""
frmViscIng.Combo1.Text = ""
frmViscIng.Text2.Text = ""
frmViscIng.Text3.Text = ""
frmViscIng.Text4.Text = ""
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Compuesto From Viscosidades Group by Compuesto")
If rs.RecordCount <> 0 Then
rs.MoveFirst
Do Until rs.EOF = True
frmViscIng.Combo1.AddItem (rs.Fields("compuesto"))
rs.MoveNext
Loop
End If
db.Close
frmViscIng.Visible = True
End Sub

Private Sub mnuingvischist_Click()
Form1.Enabled = False
Form1.Visible = False
frmViscIng.Show
frmViscIng.Visible = False
frmViscIng.Combo1.Text = ""
frmViscIng.Text2.Text = ""
frmViscIng.Text3.Text = ""
frmViscIng.Text4.Text = ""
frmViscIng.Text5.Enabled = True
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Compuesto From Viscosidades Group by Compuesto")
If rs.RecordCount <> 0 Then
rs.MoveFirst
Do Until rs.EOF = True
frmViscIng.Combo1.AddItem (rs.Fields("compuesto"))
rs.MoveNext
Loop
End If
db.Close
frmViscIng.Visible = True
frmViscIng.Text5.Text = ""
frmViscIng.Text5.SetFocus
End Sub

Private Sub mnuIny_Click()
Form1.Enabled = False
Form1.Visible = False
frmIny.Show
End Sub
Private Sub mnuLabDur_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
If Form1.Caption <> "Entorno Bafir - Entorno de Gestión de planta - Modo Laboratorio" Then
Dim db1 As Database
Dim rs1 As Recordset
Set db1 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db1.OpenRecordset("REG", dbOpenTable)
rs1.Index = "primarykey"
rs1.Seek "=", "LabKey"
strPass = rs1.Fields("dato")
ingr = InputBox("Ingrese la clave", "Clave de ingreso")
If ingr <> strPass Then
    roro = MsgBox("La clave no es valida", vbCritical + vbOKOnly, "Error")
    db1.Close
    Exit Sub
End If
Else
frmCargarDureza.Combo1.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select tipo_ensayo From durezas_produccion Group by tipo_ensayo")
Do Until rs.EOF = True
frmCargarDureza.Combo1.AddItem (rs.Fields("tipo_ensayo"))
rs.MoveNext
Loop
db.Close




''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
Form1.Enabled = False
frmCargarDureza.Show
Form1.Visible = False
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
End If
''''''''''asi era el ingreso viejo por clave de laboratorio sacar cuando se active login
End Sub
Private Sub mnuListado_Click()
FrmLotedureza.Show
Form1.Enabled = False
Form1.Visible = False
End Sub

Private Sub mnuLote_Click()
frmAEDLOTES.Text1.Text = ""
frmAEDLOTES.Text2.Text = ""
frmAEDLOTES.Text3.Text = ""
frmAEDLOTES.Text4.Text = ""
frmAEDLOTES.Check1.Value = False
frmAEDLOTES.Check2.Value = False
frmAEDLOTES.Command3.Enabled = False
frmAEDLOTES.Command4.Enabled = False
Me.Enabled = False
frmAEDLOTES.Show
End Sub

Private Sub mnuLotesPegassus_Click()
ShellExecute 0&, vbNullString, "\\Servidor2\e\EntornoBafir\planillas\Carga de lotes en pegassus.xls", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub mnumantpanel_Click()
Call confreg
If confr = "punto" Then
    dasdasd = MsgBox("El programa está diseñado para funcionar con la convención de signos del pais, en la cual se fija al 'punto' como separador decimal, y a la 'coma' como separador de miles. Otra configuración hará que la carga de los valores sea errónea. La unica sección que se ha diseñado para funcionar bajo esta convención no standard para el Pais, es la sección de administración", vbCritical + vbOKOnly, "Atención")
    Exit Sub
End If
a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
rcset.Open "SELECT Nombre FROM Mantenimiento_Operarios", CONN, adOpenStatic, adLockReadOnly
frmUsuarioYContr.Combo1.Clear
frmUsuarioYContr.Text2.Text = ""
Do Until rcset.EOF = True
    frmUsuarioYContr.Combo1.AddItem (rcset.Fields("nombre"))
    rcset.MoveNext
Loop
rcset.Close

frmUsuarioYContr.Show (1)


rcset.Open "SELECT password,Auth FROM Mantenimiento_Operarios where nombre = '" & frmUsuarioYContr.usuario & "'", CONN, adOpenStatic, adLockReadOnly
If rcset.RecordCount = 0 Then
    sdfgdfgg = MsgBox("Los datos ingresados son incorrectos", vbCritical + vbOKOnly, "Error")
    CONN.Close
    Exit Sub
Else
    If frmUsuarioYContr.contrasena <> rcset.Fields("password") Then
        sdfgdfgg = MsgBox("Los datos ingresados son incorrectos", vbCritical + vbOKOnly, "Error")
        CONN.Close
        Exit Sub
    End If
End If
    If rcset.Fields("Auth") = 1 Then
        Me.Enabled = False
        Me.Visible = False
        frmMantCargar.Label7.Caption = frmUsuarioYContr.usuario
        frmMantCargar.traedaTos
        frmmantpanel.Show
    Else
    
        Me.Enabled = False
        Me.Visible = False
        formUlario = Me.Name
        frmMantCargar.Label7.Caption = frmUsuarioYContr.usuario
        frmMantCargar.traedaTos
        frmMantCargar.Show
    End If






CONN.Close
End Sub

Private Sub mnuManual_Click()
ShellExecute 0&, vbNullString, "\\Servidor2\e\EntornoBafir\Documentos\Manual Entorno Bafir.PDF", vbNullString, vbNullString, vbMaximizedFocus
End Sub
Private Sub mnuManualCalidad_Click()
ShellExecute 0&, vbNullString, "\\Servidor2\e\ISO9000\MANUAL~1\manual~2.doc", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnumedicionespegasus_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

desdelote = InputBox("Ingrese el lote DESDE el que quiere comenzar a actualizar las mediciones")

sPathBase = "\\Servidor2\e\laboratorio\Pegassus\IntelligentWorld\Datos\base.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT * FROM lotes where ID >= " & desdelote, cnn, adOpenStatic, adLockReadOnly
    contador = 0
    Do Until rst.EOF = True
        rst1.Open "SELECT * FROM ensayodimensional where IDLOTE = " & rst.Fields("ID"), cnn, adOpenStatic, adLockOptimistic
        If rst1.RecordCount = 0 Then
            rst1.AddNew
            rst1.Fields("IDCOTA") = "9"
            rst1.Fields("valor") = "OK"
            rst1.Fields("idinstrumento") = "1"
            rst1.Fields("IDLOTE") = rst.Fields("ID")
            rst1.Update
            contador = contador + 1
        End If
        rst1.Close
        rst.MoveNext
    Loop
    MsgBox ("Se han realizado " & contador & " ingresos dimensionales en la base del Pegassus")
End Sub

Private Sub mnuMezcla_Click()
Me.Enabled = False

frmAEDMEZCLA.Combo1.Clear
frmAEDMEZCLA.Text1.Text = ""

frmAEDMEZCLA.Text3.Text = ""
frmAEDMEZCLA.Text4.Text = ""
frmAEDMEZCLA.Check1.Value = False
frmAEDMEZCLA.Check2.Value = False
frmAEDMEZCLA.Check3.Value = False
frmAEDMEZCLA.Command3.Enabled = False
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
Set rst1 = New ADODB.Recordset

sPathBase = "\\Servidor2\e\entornobafir\AED.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With
    'rst simulacion
    'rst1 real
    
    rst.Open "SELECT batch FROM mezclas order by batch", cnn, adOpenStatic, adLockReadOnly
    If rst.RecordCount = 0 Then
    
    Else
    Do Until rst.EOF = True
        frmAEDMEZCLA.Combo1.AddItem (rst.Fields("batch"))
        rst.MoveNext
    Loop
    
    End If
    rst.Close
    rst.Open "SELECT * FROM formbase where cualidad = 'AED'", cnn, adOpenStatic, adLockReadOnly
    frmAEDMEZCLA.Text3.Clear
    Do Until rst.EOF = True
        frmAEDMEZCLA.Text3.AddItem (rst.Fields("N_FORMULA"))
        a = rst.Fields("cualidad")
        rst.MoveNext
    Loop
    
cnn.Close
frmAEDMEZCLA.Show
End Sub

Private Sub mnuMezclaAprob_Click()
Form1.Enabled = False
Form1.Visible = False
frmMezclasAprob.Show
frmMezclasAprob.MSFlexGrid1.Clear
End Sub

Private Sub mnuMezcladoPend_Click()
Form1.Enabled = False
Form1.Visible = False
frmMezcladoPend.Show
frmMezcladoPend.Command3.Enabled = False
End Sub

Private Sub mnumezclareg_Click()
frmmezclareg.Combo1.Clear
frmmezclareg.Combo2.Clear
frmmezclareg.Combo3.Clear
frmmezclareg.MSFlexGrid1.Clear
frmmezclareg.Text1.Text = ""
frmmezclareg.Text2.Text = ""
frmmezclareg.Text3.Text = ""
frmmezclareg.Check1.Value = 0
frmmezclareg.Check2.Value = 0

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select compuesto from pesado group by compuesto")

rs.MoveFirst

Do Until rs.EOF = True
    frmmezclareg.Combo1.AddItem (rs.Fields("compuesto"))
    rs.MoveNext
Loop
Form1.Enabled = False
Form1.Visible = False
frmmezclareg.Show
db.Close
frmmezclareg.Combo1.SetFocus
End Sub

Private Sub mnuMP_Click()
frmMP.MSFlexGrid1.Clear
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT cod_prod,descrip FROM producto where apto = true order by descrip", cnn, adOpenStatic, adLockReadOnly
    
    frmMP.MSFlexGrid1.Rows = rst.RecordCount + 1
    frmMP.MSFlexGrid1.Cols = 2
    frmMP.MSFlexGrid1.TextMatrix(0, 0) = "Codigo"
    frmMP.MSFlexGrid1.TextMatrix(0, 1) = "Nombre"
    
        Fila = 1
    Do Until rst.EOF = True
        frmMP.MSFlexGrid1.TextMatrix(Fila, 0) = rst.Fields("cod_prod")
        frmMP.MSFlexGrid1.TextMatrix(Fila, 1) = rst.Fields("descrip")
        Fila = Fila + 1
        rst.MoveNext
    Loop
    AutoGrid frmMP.MSFlexGrid1, 2
    cnn.Close
    frmMP.Show
    frmMP.MSFlexGrid1.ColWidth(0) = 1000
    frmMP.MSFlexGrid1.ColWidth(1) = 5000
    
    
End Sub

Private Sub mnuNormaAsign_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT N_formula FROM formbase order by N_formula", cnn, adOpenStatic, adLockReadOnly
frmNormaAsign.List1.Clear
frmNormaAsign.List2.Clear
frmNormaAsign.List3.Clear
frmNormaAsign.List4.Clear
frmNormaAsign.List5.Clear
frmNormaAsign.List6.Clear
frmNormaAsign.List7.Clear

Do Until rst.EOF = True
    frmNormaAsign.List1.AddItem (rst.Fields("N_formula")) & ""
    rst.MoveNext
Loop

rst.Close
rst.Open "SELECT Norma FROM Norma order by norma", cnn, adOpenStatic, adLockReadOnly

If rst.RecordCount <> 0 Then
    Do Until rst.EOF = True
        frmNormaAsign.List2.AddItem (rst.Fields("Norma"))
        rst.MoveNext
    Loop
End If
rst.Close
rst.Open "SELECT Cliente FROM clientes", cnn, adOpenStatic, adLockReadOnly
'frmNormaAsign.List4.AddItem "Todos"
Do Until rst.EOF = True
    'frmNormaAsign.List4.AddItem (rst.Fields("Cliente"))
    frmNormaAsign.List5.AddItem (rst.Fields("Cliente"))
    rst.MoveNext
Loop

rst.Close
rst.Open "SELECT pieza FROM pieza", cnn, adOpenStatic, adLockReadOnly
'frmNormaAsign.List4.AddItem "Todos"
frmNormaAsign.List6.AddItem ("Sin Asignar")
Do Until rst.EOF = True
    'frmNormaAsign.List4.AddItem (rst.Fields("Cliente"))
    frmNormaAsign.List6.AddItem (rst.Fields("pieza"))
    rst.MoveNext
Loop






'frmNormaAsign.List4.Text = "Todos"
frmNormaAsign.List5.Text = "Bafir"
cnn.Close




Me.Enabled = False
frmNormaAsign.Show
End Sub

Private Sub mnuNormas_Click()
Form1.Enabled = False
Form1.Visible = False

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Norma from normas")
Do Until rs.EOF = True
frmNormas.Combo1.AddItem (rs.Fields("Norma"))
rs.MoveNext
Loop
db.Close
frmNormas.Command6.Visible = False
frmNormas.Command7.Visible = False
frmNormas.Command8.Visible = False
frmNormas.Command3.Enabled = False
frmNormas.Command4.Enabled = False
frmNormas.Combo1.Text = ""
frmNormas.Text1.Text = ""
frmNormas.Text2.Text = ""
frmNormas.Text3.Text = ""
frmNormas.Text4.Text = ""
frmNormas.Text1.Enabled = False
frmNormas.Text2.Enabled = False
frmNormas.Text3.Enabled = False
frmNormas.Text4.Enabled = False
frmNormas.Show
End Sub

Private Sub mnuOrganigrama_Click()
ShellExecute 0&, vbNullString, "\\Servidor2\e\ISO9000\MANUAL DE CALIDAD\Organigrama.doc", vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub mnuPartidasNuevas_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
frmPartidasNuevas.MSFlexGrid1.Clear
frmPartidasNuevas.Command8.Visible = True
frmPartidasNuevas.Show
Form1.Enabled = False
Form1.Visible = False
End Sub

Private Sub mnuPCPaltapieza_Click()
Form1.Enabled = False
frmPCPaltapieza.Show
End Sub

Private Sub mnuPegassus_Click()
Dim strFic As String
Dim strParam As String

strFic = "\\Servidor2\E\EntornoBafir\Pegassus MOD\Pegassus.exe"


Shell strFic, vbNormalFocus

aasd = MsgBox("Se está ejecutando el Pegassus Version Modificada", vbInformation + vbOKOnly, "Pegassus")

End Sub

Private Sub mnupiezasactivas_Click()
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT * FROM piezas_en_produccion", cnn, adOpenStatic, adLockReadOnly
    If rst.RecordCount = 0 Then
        sdfsf = MsgBox("No hay piezas activas", vbCritical + vbOKOnly, "Error")
        cnn.Close
        Exit Sub
    End If
    frmPiezasActivas.MSFlexGrid1.Rows = rst.RecordCount + 1
    frmPiezasActivas.MSFlexGrid1.TextMatrix(0, 0) = "Pieza"
    frmPiezasActivas.MSFlexGrid1.TextMatrix(0, 1) = "Cliente"
    frmPiezasActivas.MSFlexGrid1.TextMatrix(0, 2) = "Norma"
    frmPiezasActivas.MSFlexGrid1.TextMatrix(0, 3) = "Compuesto"
    frmPiezasActivas.MSFlexGrid1.TextMatrix(0, 4) = "Fecha"
    Fila = 1
    Do Until rst.EOF = True
        frmPiezasActivas.MSFlexGrid1.TextMatrix(Fila, 0) = rst.Fields("pieza")
        frmPiezasActivas.MSFlexGrid1.TextMatrix(Fila, 1) = rst.Fields("Cliente")
        frmPiezasActivas.MSFlexGrid1.TextMatrix(Fila, 2) = rst.Fields("norma")
        frmPiezasActivas.MSFlexGrid1.TextMatrix(Fila, 3) = rst.Fields("compuesto")
        frmPiezasActivas.MSFlexGrid1.TextMatrix(Fila, 4) = rst.Fields("fecha_alta")
        Fila = Fila + 1
        rst.MoveNext
    Loop
    cnn.Close
    AutoGrid frmPiezasActivas.MSFlexGrid1, 5
    Me.Enabled = False
    frmPiezasActivas.Show
End Sub

Private Sub mnuRecBuscaComp_Click()
frmPartidasNuevas.MSFlexGrid1.Clear
frmPartidasNuevas.Command8.Visible = False
frmPartidasNuevas.Show
Form1.Enabled = False
Form1.Visible = False
End Sub

Private Sub mnuRecComp_Click()
frmRecComp.Show
Form1.Enabled = False
Form1.Visible = False
frmRecComp.Text1 = ""
frmRecComp.Text2 = ""
frmRecComp.Text3 = ""
frmRecComp.Text1.SetFocus
End Sub

Private Sub mnuReci_Click()
Form1.Enabled = False
Form1.Visible = False
frmCuerdas.Show
frmCuerdas.Text1.Text = ""
frmCuerdas.Combo3.Clear
frmCuerdas.Text2.Text = ""
frmCuerdas.Text3.Text = ""
frmCuerdas.Text4.Text = ""
frmCuerdas.Combo2.Clear
frmCuerdas.Combo1.Clear
frmCuerdas.Check1.Value = 0
frmCuerdas.Combo4.Clear
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select proveedor from cuerdas group by proveedor")
Set rs1 = db.OpenRecordset("Select cuerda from cuerdas group by cuerda")
Set rs2 = db.OpenRecordset("Select material from cuerdas group by material")
Set rs3 = db.OpenRecordset("Select responsable from cuerdas group by responsable")
Do Until rs.EOF = True
frmCuerdas.Combo3.AddItem (rs.Fields("proveedor"))
rs.MoveNext
Loop
Do Until rs1.EOF = True
frmCuerdas.Combo2.AddItem (rs1.Fields("cuerda"))
rs1.MoveNext
Loop
Do Until rs2.EOF = True
frmCuerdas.Combo1.AddItem (rs2.Fields("material"))
rs2.MoveNext
Loop
Do Until rs3.EOF = True
frmCuerdas.Combo4.AddItem (rs3.Fields("responsable"))
rs3.MoveNext
Loop
frmCuerdas.Text1 = Date
frmCuerdas.Text4 = Format(Date, "YYMMDD")
db.Close
End Sub
Private Sub mnuRecom_Click()
If Form1.Caption <> "Entorno Bafir - Entorno de Gestión de planta - Modo Laboratorio" Then
Dim db1 As Database
Dim rs1 As Recordset
Set db1 = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db1.OpenRecordset("REG", dbOpenTable)
rs1.Index = "primarykey"
rs1.Seek "=", "LabKey"
strPass = rs1.Fields("dato")
ingr = InputBox("Ingrese la clave", "Clave de ingreso")
If ingr <> strPass Then
    roro = MsgBox("La clave no es valida", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Else
frmRecom.Show
Form1.Visible = False
Form1.Enabled = False
frmRecom.Text5.TabIndex = 1
frmRecom.Label4.TabIndex = 2
frmRecom.Label10.TabIndex = 3
frmRecom.Label11.TabIndex = 4
frmRecom.Label12.TabIndex = 5
frmRecom.Label13.TabIndex = 6
frmRecom.Text7.TabIndex = 7
frmRecom.Text1.TabIndex = 8
frmRecom.Text2.TabIndex = 9
frmRecom.Text3.TabIndex = 10
frmRecom.Label18.TabIndex = 11
frmRecom.Text4.TabIndex = 12
frmRecom.Text6.TabIndex = 13
frmRecom.Text5.Text = ""
frmRecom.Label4.Caption = ""
frmRecom.Label10.Caption = ""
frmRecom.Label11.Caption = ""
frmRecom.Label12.Caption = ""
frmRecom.Label13.Caption = ""
frmRecom.Text7.Text = ""
frmRecom.Text1.Text = ""
frmRecom.Text1.Enabled = False
frmRecom.Text2.Text = ""
frmRecom.Text2.Enabled = False
frmRecom.Text3.Text = ""
frmRecom.Text3.Enabled = False
frmRecom.Label18.Caption = ""
frmRecom.Text4.Text = ""
frmRecom.Text4.Enabled = False
frmRecom.Text6.Text = ""
frmRecom.Text6.Enabled = False
frmRecom.Command4.Enabled = False
frmRecom.Command3.Enabled = True
frmRecom.Text5.Enabled = True
frmRecom.Text5.SetFocus
End If
End Sub
Private Sub mnuSalir_Click()
End
End Sub

Private Sub mnuReometro_Click()
frmReometro.List1.Clear
frmReometro.List2.Clear
frmReometroCompSel.Combo1.Clear
Dim CONN As New ADODB.Connection
Dim rs As New ADODB.Recordset

CONN.ConnectionString = "Driver={Microsoft Visual FoxPro Driver};" & "SourceType=DBF;" & "SourceDB=F:\laboratorio\reometro\backup reotron\051227\;" & "Exclusive=No;"
CONN.Open
Set rs = CONN.Execute("SELECT comp FROM reod group by comp;")
Do Until rs.EOF = True
frmReometro.List2.AddItem (rs.Fields("comp"))
frmReometroCompSel.Combo1.AddItem (rs.Fields("comp"))
rs.MoveNext
Loop

CONN.Close

Dim db As Database
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db.OpenRecordset("Select compuesto from reomcalcular")

Do Until rs1.EOF = True
frmReometro.List1.AddItem (rs1.Fields("compuesto"))
'frmReometro.List2.re rs1.Fields("compuesto")
rs1.MoveNext
Loop

db.Close


Form1.Enabled = False
Form1.Visible = False
frmReometro.Show
End Sub

Private Sub mnuReometroNuevo_Click()
Me.Enabled = False
Me.Visible = False


'''''''''
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset

Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset

sPathBase = "\\REOMETRO\tisa\reo.mdb"

With cnn
     'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
     .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";"
End With
Reintentar:
     On Error Resume Next
     cnn.Open
    If Err.Number <> 0 Then
        If Err.Number = -2147467259 Then
            dasdasd = MsgBox("La base del reómetro está siendo utilizada por el programa de trazado y no se puede acceder a la misma. Desea reintentar acceder a la misma?.", vbInformation + vbYesNo, "Base Ocupada")
            If dasdasd = vbYes Then
                reit = True
            Else
                frmReometroNuevo.Enabled = True
                Exit Sub
            End If
        Else
            e = Err.Number
            d = Err.Description
            asdsdf = MsgBox("Error No previsto. Codigo " & Err.Number & " Descripción " & Err.Description, vbCritical + vbOKOnly, "Error no previsto")
            Exit Sub
        End If
        If reit = True Then
            GoTo Reintentar
        End If
    End If
    On Error GoTo 0


rst.Open "SELECT Compound FROM Parametertable order by compound asc", cnn, adOpenStatic, adLockReadOnly
frmReometroNuevo.Combo1.Clear
frmReometroNuevo.Combo2.Clear
frmReometroNuevo.Combo3.Clear
frmReometroNuevo.Combo4.Clear
frmReometroNuevo.MSFlexGrid1.Clear
frmReometroNuevo.MSFlexGrid1.Cols = 9
frmReometroNuevo.MSFlexGrid1.Rows = 1

Do Until rst.EOF = True
    frmReometroNuevo.Combo1.AddItem (rst.Fields("compound"))
    rst.MoveNext
Loop
cnn.Close
frmReometroNuevo.Show
End Sub

Private Sub mnuSol_Click()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select codigo_recomendacion from compuestos_para_cotizacion")
Set rs1 = db.OpenRecordset("Select Cliente from compuestos_para_cotizacion group by cliente")
Set rs2 = db.OpenRecordset("Select respons_solici from compuestos_para_cotizacion group by respons_solici")
If rs.RecordCount = 0 Then
codigorecoUltimo = 1
Else
rs.MoveLast
codigorecoUltimo = rs.Fields("codigo_recomendacion") + 1
End If
frmSolcoti.Label2 = codigorecoUltimo
frmSolcoti.Label4 = Date
frmSolcoti.Text1.Text = ""
frmSolcoti.Text1.TabIndex = 1
frmSolcoti.Text2.TabIndex = 2
frmSolcoti.Text3.TabIndex = 3
frmSolcoti.Combo1.TabIndex = 4
frmSolcoti.Text5.TabIndex = 5
frmSolcoti.Text2.Text = ""
frmSolcoti.Text3.Text = ""
frmSolcoti.Text5.Text = ""
frmSolcoti.Combo1.Text = ""
frmSolcoti.Combo1.Clear
rs.Close
If rs1.RecordCount > 0 Then
rs1.MoveFirst
Do While rs1.EOF = False
frmSolcoti.Combo1.AddItem (rs1.Fields("cliente"))
rs1.MoveNext
Loop
rs1.Close
End If
rs2.MoveFirst
Do While rs2.EOF = False
frmSolcoti.Text1.AddItem (rs2.Fields("respons_solici"))
rs2.MoveNext
Loop
Form1.Visible = False
Form1.Enabled = False
frmSolcoti.Show
db.Close
End Sub
Private Sub mnuSolicitar_Click()
Form1.Enabled = False
Form1.Visible = False
frmDureza.Show
End Sub

Private Sub mnuStock_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Form1.Enabled = False
Form1.Visible = False
frmStock.Controla_minimo
frmStock.Show
End Sub

Private Sub mnuTermocupla_Click()
Form1.Visible = False
Form1.Enabled = False
frmTermocupla.Show
End Sub

Private Sub mnuTiempComun_Click()
frmTiempoComun.Command1.Enabled = False
frmTiempoComun.Command2.Enabled = False
frmTiempoComun.Command3.Enabled = True
frmTiempoComun.Command4.Enabled = True
frmTiempoComun.Command5.Enabled = False
frmTiempoComun.Command6.Enabled = True
frmTiempoComun.Command7.Enabled = False
frmTiempoComun.Label8.Caption = ""
frmTiempoComun.Label9.Caption = ""
frmTiempoComun.Combo1.Clear
frmTiempoComun.Combo2.Clear
frmTiempoComun.Text3.Text = ""
frmTiempoComun.Text4.Text = ""
frmTiempoComun.Text5.Text = ""
frmTiempoComun.Text1.Text = ""
frmTiempoComun.Text2.Text = ""
frmTiempoComun.Text6.Text = ""
frmTiempoComun.Text7.Text = ""
frmTiempoComun.Text8.Text = ""
frmTiempoComun.Text9.Text = ""
frmTiempoComun.Show

Me.Enabled = False
End Sub

Private Sub mnuTiempoAed_Click()
Me.Enabled = False
frmTiempoAed.Show
End Sub

Private Sub mnuTolerancia_Click()
frmViscTol.Show
Form1.Enabled = False
Form1.Visible = False
frmViscTol.Combo1.Text = ""
frmViscTol.Text1.Text = ""
frmViscTol.Text2.Text = ""
frmViscTol.Text3.Text = ""
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select Compuesto from viscosidades group by compuesto")
If rs.RecordCount = 0 Then
sdfdf = MsgBox("No hay registros", vbCritical + vbOKOnly, "Error")
Exit Sub
End If
Do Until rs.EOF = True
frmViscTol.Combo1.AddItem (rs.Fields("Compuesto"))
rs.MoveNext
Loop
End Sub

Private Sub mnutraccionIndividual_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
Form1.Enabled = False
Form1.Visible = False
frmTraccionIndividual.Show
End Sub

Private Sub mnuValtracc_Click()
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Do
    cantidadensayos = InputBox("Ingrese la cantidad de partidas de las cuales ingresará sus datos de tracción", "Cantidad de partidas")
    If IsNumeric(cantidadensayos) = False Then
        gf = MsgBox("Debe ingresar un número entero válido", vbCritical + vbOKOnly, "Error")
    End If
    If cantidadensayos = "" Then
        Exit Sub
    End If
Loop Until IsNumeric(cantidadensayos) = True
cantidadensayos = CInt(cantidadensayos)
ReDim codigoensayos(1 To cantidadensayos, 1, 1, 10, 10, 1)
For bucleensayos = 1 To cantidadensayos 'este es el que se repite una vez por cada ensayo y solo obtiene los codigos de los ensayos y la cantidad de especimenes de cada uno
    
    Do
        codigoensayos(bucleensayos, 0, 0, 0, 0, 0) = InputBox("Ingrese el codigo de ensayo", "Código de ensayo " & bucleensayos & "/" & cantidadensayos)
        If codigoensayos(bucleensayos, 0, 0, 0, 0, 0) = "" Then
            Exit Sub
        End If
        If codigoensayos(bucleensayos, 0, 0, 0, 0, 0) = "" Then
            Exit Sub
        End If
        Set rs = db.OpenRecordset("Select Codigo_ensayo, compuesto, partida from traccion where codigo_ensayo =" & codigoensayos(bucleensayos, 0, 0, 0, 0, 0))
        Set rs1 = db.OpenRecordset("Select codigo_ensayo, especimenes from traccion_dimensiones where codigo_ensayo =" & codigoensayos(bucleensayos, 0, 0, 0, 0, 0))
        If rs.RecordCount = 1 And rs1.RecordCount = 0 Then
        sdff = MsgBox("La partida ya ha sido cargada", vbCritical + vbOKOnly, "Partida Cerrada")
        Exit Sub
        End If
'        MsgBox ("Select Codigo_ensayo, especimenes, compuesto, partida from traccion where codigo_ensayo ='" & codigoensayos(bucleensayos, 0, 0, 0, 0, 0) & ";'")
        If rs.RecordCount = 0 Then
            sdfgsdg = MsgBox("No existe el codigo ingresado", vbCritical + vbOKOnly, "Error")
        End If
    Loop Until RecordCount >= 0
    
    codigoensayos(bucleensayos, 1, 0, 0, 0, 0) = rs1.Fields("especimenes")
    codigoensayos(bucleensayos, 0, 1, 0, 0, 0) = rs.Fields("compuesto") & "-" & rs.Fields("partida")
Next 'aca termina bucle ensayos
Volver2:
For bucleobtencion = 1 To cantidadensayos 'este bucle obtiene los datos de tracción y elongación de cada. Es el total de los ensayos, no es el bucle de probetas (especimenes)
    For bucleespecimenes = 1 To codigoensayos(bucleobtencion, 1, 0, 0, 0, 0) ' este bucle es para cada probeta (especimenes)

        codigoensayos(bucleobtencion, 0, 0, bucleespecimenes, 0, 0) = InputBox("Ingrese la tracción " & bucleespecimenes & "/" & codigoensayos(bucleobtencion, 1, 0, 0, 0, 0), codigoensayos(bucleobtencion, 0, 1, 0, 0, 0)) 'ingreso de traccion
        If codigoensayos(bucleobtencion, 0, 0, bucleespecimenes, 0, 0) = "" Then
            GoTo Volver2
        End If
        punto = InStr(1, codigoensayos(bucleobtencion, 0, 0, bucleespecimenes, 0, 0), ".")
        If punto <> 0 Then
        Mid(codigoensayos(bucleobtencion, 0, 0, bucleespecimenes, 0, 0), punto) = ","
        End If
    Next ' bucleespecimenes
    For bucleespecimenes1 = 1 To codigoensayos(bucleobtencion, 1, 0, 0, 0, 0) ' este bucle es para cada probeta (especimenes1)
volver3:
        codigoensayos(bucleobtencion, 0, 0, 0, bucleespecimenes1, 0) = InputBox("Ingrese la Elongación " & bucleespecimenes1 & "/" & codigoensayos(bucleobtencion, 1, 0, 0, 0, 0), codigoensayos(bucleobtencion, 0, 1, 0, 0, 0))    'ingreso de elongación
        If codigoensayos(bucleobtencion, 0, 0, 0, bucleespecimenes1, 0) = "" Then
        GoTo volver3
        End If
        punto1 = InStr(1, codigoensayos(bucleobtencion, 0, 0, 0, 1, 0), ".")
        If punto1 <> 0 Then
        Mid(codigoensayos(bucleobtencion, 0, 0, 0, 1, 0), punto1) = ","
        End If
    Next 'de bucle especimenes1
    'codigoensayos(bucleobtencion, 0, 0, 0, 0, 1) = InputBox("Ingrese la dureza del compuesto " & codigoensayos(bucleobtencion, 0, 1, 0, 0, 0), "Dureza") ' dureza
Next ' de bucleobtencion
'Aca finaliza la parte de ingreso de datos
'aca comienza la parte de calculo
'acordate que tenes que hacer el ingreso de la separación de la medicion de elongacion 25 o 20
For buclecalculo = 1 To cantidadensayos 'bucle de cada ensayos
    traccion = 0
    elongacion = 0
    For buclecalculoespecimen = 1 To codigoensayos(buclecalculo, 1, 0, 0, 0, 0) 'bucle para especimenes
        traccion = traccion + codigoensayos(buclecalculo, 0, 0, buclecalculoespecimen, 0, 0)
        elongacion = elongacion + codigoensayos(buclecalculo, 0, 0, 0, buclecalculoespecimen, 0)
    Next 'buclecalculoespecimen
    traccionpromedio = traccion / codigoensayos(buclecalculo, 1, 0, 0, 0, 0)
    elongacionpromedio = elongacion / codigoensayos(buclecalculo, 1, 0, 0, 0, 0)
    Set rs = db.OpenRecordset("Select CODIGO_ENSAYO, COMPUESTO, REFERENCIA, PARTIDA, ESTADO_ENSAYO, DUREZA, TRACCION, ELONGACION from traccion where CODIGO_ENSAYO = " & codigoensayos(buclecalculo, 0, 0, 0, 0, 0))
    Set rs1 = db.OpenRecordset("Select codigo_ensayo, especimenes, espesor, ancho from traccion_dimensiones where codigo_ensayo = " & codigoensayos(buclecalculo, 0, 0, 0, 0, 0))
    ReDim inicioespesor(1 To (codigoensayos(buclecalculo, 1, 0, 0, 0, 0)))
    ReDim inicioancho(1 To (codigoensayos(buclecalculo, 1, 0, 0, 0, 0)))
    ReDim espesor(1 To codigoensayos(buclecalculo, 1, 0, 0, 0, 0))
    ReDim ancho(1 To codigoensayos(buclecalculo, 1, 0, 0, 0, 0))
    inicioespesor(1) = 1
    inicioancho(1) = 1
    totalesp = Len(rs1.Fields("espesor"))
    totalanch = Len(rs1.Fields("ancho"))
    For bucleespesor = 1 To (codigoensayos(buclecalculo, 1, 0, 0, 0, 0) - 1)
        inicioespesor(bucleespesor + 1) = InStr(inicioespesor(bucleespesor) + 1, rs1.Fields("espesor"), "@", vbTextCompare)
    Next ' bucleespesor
    For bucleancho = 1 To (codigoensayos(buclecalculo, 1, 0, 0, 0, 0) - 1)
        inicioancho(bucleancho + 1) = InStr(inicioancho(bucleancho) + 1, rs1.Fields("ancho"), "@", vbTextCompare)
    Next ' bucleancho
    For bucleextraccion = 1 To codigoensayos(buclecalculo, 1, 0, 0, 0, 0)
        If bucleextraccion = 1 Then
            espesor(bucleextraccion) = Mid(rs1.Fields("espesor"), inicioespesor(bucleextraccion), inicioespesor(bucleextraccion + 1) - 1)
            ancho(bucleextraccion) = Mid(rs1.Fields("ancho"), inicioancho(bucleextraccion), inicioancho(bucleextraccion + 1) - 1)
        End If ' If bucleextraccion = 1 Then
        If bucleextraccion > 1 And bucleextraccion <> CInt(codigoensayos(buclecalculo, 1, 0, 0, 0, 0)) Then
            espesor(bucleextraccion) = Mid(rs1.Fields("espesor"), inicioespesor(bucleextraccion) + 1, (inicioespesor(bucleextraccion + 1) - 1) - inicioespesor(bucleextraccion))
            ancho(bucleextraccion) = Mid(rs1.Fields("ancho"), inicioancho(bucleextraccion) + 1, (inicioancho(bucleextraccion + 1) - 1) - inicioancho(bucleextraccion))
        End If ' bucleextraccion > 1 And bucleextraccion <> codigoensayos(buclecalculo, 1, 0, 0, 0, 0) Then
        If bucleextraccion = CInt(codigoensayos(buclecalculo, 1, 0, 0, 0, 0)) Then
            espesor(bucleextraccion) = Right(rs1.Fields("espesor"), totalesp - (inicioespesor(bucleextraccion)))
            ancho(bucleextraccion) = Right(rs1.Fields("ancho"), totalanch - (inicioancho(bucleextraccion)))
        End If 'bucleextraccion = codigoensayos(buclecalculo, 1, 0, 0, 0, 0) Then
   Next 'bucleextraccion
   promedioespesor = 0
   promedioancho = 0
   For buclepromedio = 1 To codigoensayos(buclecalculo, 1, 0, 0, 0, 0)
        promedioespesor = promedioespesor + espesor(buclepromedio)
        promedioancho = promedioancho + ancho(buclepromedio)
   Next ' buclepromedio
   promedioespesorfinal = promedioespesor / codigoensayos(buclecalculo, 1, 0, 0, 0, 0)
   promedioanchofinal = promedioancho / codigoensayos(buclecalculo, 1, 0, 0, 0, 0)
   traccion = (traccionpromedio / (promedioanchofinal * promedioespesorfinal)) * 9.81 ' no detecto el punto en la traccion promedio
   separacion = InputBox("Ingrese la longitud de la separación de las marcas para la medición de la elongación en la probeta", "Elongación " & codigoensayos(buclecalculo, 0, 1, 0, 0, 0))
   elongacion = 100 * ((elongacionpromedio - separacion) / separacion)
   pant = MsgBox("Compuesto " & codigoensayos(buclecalculo, 0, 1, 0, 0, 0) & " " & rs.Fields("estado_ensayo") & " Tracción: " & traccion & " Elongación: " & elongacion, vbOKOnly + vbInformation, codigoensayos(buclecalculo, 0, 1, 0, 0, 0) & "-" & rs.Fields("estado_ensayo") & " " & rs.Fields("referencia"))
   rs.Edit
   rs.Fields("traccion") = Format(traccion, "0.000")
   rs.Fields("elongacion") = Format(elongacion, "0.0")
   rs.Update
Next ' buclecalculo (este es el que calcula los datos)
db.Close
End Sub

Private Sub mnuVerFormula_Click()
Call confreg
If confr <> "coma" Then
    dasdasd = MsgBox("Esta sección solo funciona correctamente cuando en la configuración regional está seteado el 'punto' como separador de miles y la 'coma' como separador decimal", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
    frmDesarrSelecc.Combo1.Clear
    Dim db As Database
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
    Set rs = db.OpenRecordset("Select N_Formula from partes_desarrollo group by N_formula")
    
    Do Until rs.EOF = True
        frmDesarrSelecc.Combo1.AddItem (rs.Fields("N_FORMULA"))
        rs.MoveNext
    Loop
    rs.MoveFirst
    Form1.Enabled = False
    Form1.Visible = False
    frmDesarrSelecc.Show (1)
    
    If frmDesarrSelecc.seleccionado = "Salir" Then
        Form1.Enabled = True
        Form1.Visible = True
        db.Close
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select * from partes_desarrollo where N_FORMULA = '" & frmDesarrSelecc.seleccionado & "' order by etapa asc, partes desc")
    
    rs.MoveLast
    iteMss = rs.RecordCount
    rs.MoveFirst
    
    
    '''''''''''''''''''
    
    
    
    Set rs1 = db.OpenRecordset("Select cod_prod,descrip from producto order by descrip")
    componentes = iteMss - 1
    
    Do Until rs1.EOF = True
            frmDesarrFormulaIng.Combo1(0).AddItem (rs1.Fields("descrip"))
            rs1.MoveNext
        Loop
        rs1.MoveFirst
    For i = 1 To componentes
        Load frmDesarrFormulaIng.Text1(i)
        Load frmDesarrFormulaIng.Combo1(i)
        Load frmDesarrFormulaIng.Combo2(i)
        frmDesarrFormulaIng.Text1(i).Visible = True
        Do Until rs1.EOF = True
            frmDesarrFormulaIng.Combo1(i).AddItem (rs1.Fields("descrip"))
            rs1.MoveNext
        Loop
            rs1.MoveFirst
        frmDesarrFormulaIng.Combo1(i).Visible = True
        frmDesarrFormulaIng.Combo2(i).Visible = True
        frmDesarrFormulaIng.Combo2(i).AddItem ("A")
        frmDesarrFormulaIng.Combo2(i).AddItem ("B")
        frmDesarrFormulaIng.Text1(i).Top = frmDesarrFormulaIng.Text1(i - 1).Top + frmDesarrFormulaIng.Text1(i).Height
        frmDesarrFormulaIng.Combo1(i).Top = frmDesarrFormulaIng.Combo1(i - 1).Top + frmDesarrFormulaIng.Combo1(i).Height
        frmDesarrFormulaIng.Combo2(i).Top = frmDesarrFormulaIng.Combo2(i - 1).Top + frmDesarrFormulaIng.Combo1(i).Height
    Next
    
    
    rs.MoveFirst
    For i = 0 To componentes
        
        Set rs1 = db.OpenRecordset("Select descrip from producto where cod_prod = '" & rs.Fields("cod_prod") & "'")
        frmDesarrFormulaIng.Combo1(i).Text = rs1.Fields("descrip")
        frmDesarrFormulaIng.Combo2(i).Text = rs.Fields("etapa")
        frmDesarrFormulaIng.Text1(i).Text = rs.Fields("partes")
        rs.MoveNext
    Next
    rs.MoveFirst
    frmDesarrFormulaIng.Label5.Caption = rs.Fields("N_FORMULA")
    frmDesarrFormulaIng.Label12.Caption = ""
    frmDesarrFormulaIng.Label13.Caption = ""
    frmDesarrFormulaIng.Label7.Caption = ""
    
    comOP = frmDesarrFormulaIng.Text1.Count - 1
    taBi = 0
    For i = 1 To comOP
        taBi = taBi + 3
        frmDesarrFormulaIng.Combo1(i).TabIndex = taBi
        frmDesarrFormulaIng.Text1(i).TabIndex = taBi + 1
        frmDesarrFormulaIng.Combo2(i).TabIndex = taBi + 2
    Next
    
    frmDesarrFormulaIng.Command2.Enabled = True
    frmDesarrFormulaIng.Command4.Enabled = True
    frmDesarrFormulaIng.Command5.Enabled = True
    frmDesarrFormulaIng.Command6.Enabled = True
    frmDesarrFormulaIng.Command3.Enabled = True
    frmDesarrFormulaIng.Command7.Enabled = True
    frmDesarrFormulaIng.Show
    db.Close
    
    
    
    
    
End Sub

Private Sub mnuverformulakilos_Click()
asdasd = InputBox("Ingrese Clave", "Clave")
If asdasd <> "3812" Then
    Exit Sub
End If


Me.Enabled = False
Dim db As Database
Dim rs As Recordset


Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_Formula, estado from formbase where estado = 1 or estado = 3 order by N_FORMULA")
'Set rs = db.OpenRecordset("Select N_Formula, estado from copia_formbase where estado = 1 or estado = 3 order by N_FORMULA")
'Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
'Set rs = db.OpenRecordset("Select N_formula from partes_copia group by N_formula")


frmverFormulasP.List1.Clear
frmverFormulasP.Text1.Text = ""
frmverFormulasP.MSFlexGrid1.Clear

Do Until rs.EOF = True
    frmverFormulasP.List1.AddItem (rs.Fields("N_formula"))
    rs.MoveNext
Loop


db.Close
frmverFormulasP.Show
frmverFormulasP.List1.SetFocus
End Sub

Private Sub mnuvisflujo_Click()
Form1.Enabled = False
frmLog.Show (1)
If flags = False Then
    Form1.Enabled = True
    Exit Sub
Else
    a = conectarmysql("192.168.1.40", "bafir", "bafiruser", "rafaelcapo")
    rcset.Open "SELECT clave, permisos FROM flujo_fondos_usuarios where usuario = '" & logUser & "'", CONN, adOpenStatic, adLockReadOnly
    If rcset.RecordCount = 0 Then
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
    If rcset.Fields("clave") = logPass Then
        If rcset.Fields("permisos") = 0 Then
            rcset.Close
            frmInformefondos.Show
        Else
            asd = MsgBox("Usted no posee permisos para realizar esta operación", vbCritical + vbOKOnly, "Acceso restringido")
            Form1.Enabled = True
            rcset.Close
            Exit Sub
        End If
    Else
        asd = MsgBox("Usuario o clave incorrectos", vbCritical + vbOKOnly, "Usuario o clave incorrectos")
        Form1.Enabled = True
        rcset.Close
        Exit Sub
    End If
End If
End Sub

Private Sub Timer1_Timer()
Label8.Caption = "Tiempo restante: " & reg
Timer1.Interval = 1000
Flag = 1
reg = reg - 1
    If reg = 0 Then
    End
    End If
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset


sPathBase = "\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb"

    With cnn
        'Uso OLEDB 4.0 por que el 3.51 tira error de ISAM instalable
        .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sPathBase & ";" & "Jet OLEDB:Database Password=flanflus;"
        .Open
    End With

   
    rst.Open "SELECT dato FROM reg where funcion = 'Servicio'", cnn, adOpenStatic, adLockReadOnly
    If rst.Fields("dato") = "0" Then
        End
    End If
    rst.Close
    cnn.Close
    
    
    
    
    
    
    
End Sub
'ANTIGUO METODO PARA OBTENER EL computername
'Public Function GetSettingString(hKey As Long, _
'    strPath As String, strValue As String, Optional _
'        Default As String) As String
'    Dim hCurKey As Long
'    Dim lResult As Long
'    Dim lValueType As Long
'    Dim strBuffer As String
'    Dim lDataBufferSize As Long
'    Dim intZeroPos As Integer
'    Dim lRegResult As Long
'    'Set up default value
'    If Not IsEmpty(Default) Then
'        GetSettingString = Default
'    Else
'        GetSettingString = ""
'    End If
'    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
'    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
'        lValueType, ByVal 0&, lDataBufferSize)
'    If lRegResult = ERROR_SUCCESS Then
'        If lValueType = REG_SZ Then
'            strBuffer = String(lDataBufferSize, " ")
'            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
'                ByVal strBuffer, lDataBufferSize)
'            intZeroPos = InStr(strBuffer, Chr$(0))
'            If intZeroPos > 0 Then
'                GetSettingString = Left$(strBuffer, intZeroPos - 1)
'            Else
'                GetSettingString = strBuffer
'            End If
'        End If
'    Else
'        'there is a problem
'    End If
'    lRegResult = RegCloseKey(hCurKey)
'End Function

Private Sub Timer2_Timer()

End Sub
Sub verformUlasD()
Dim db As Database
Dim rs As Recordset

kdfghkdfh = InputBox("Ingrese Clave", "Clave")
If kdfghkdfh <> "666" Then
    Exit Sub
End If

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\centralpesado.mdb", False, True, ";pwd=flanflus")
Set rs = db.OpenRecordset("Select N_formula from partes_desarrollo group by N_formula")

If rs.RecordCount = 0 Then
    asdasd = MsgBox("No hay formulas de desarrollo cargadas. saliendo", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

frmverFormulaD.List1.Clear
frmverFormulaD.Text1.Text = ""
frmverFormulaD.MSFlexGrid1.Clear

Do Until rs.EOF = True
    frmverFormulaD.List1.AddItem (rs.Fields("N_formula"))
    rs.MoveNext
Loop

Form1.Hide
db.Close

frmverFormulaD.Show
frmverFormulaD.List1.SetFocus
End Sub
