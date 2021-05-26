VERSION 5.00
Begin VB.Form frmCargarLote 
   Caption         =   "Carga Manual de Lotes"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11115
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   6570
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6120
      TabIndex        =   26
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Text            =   "Combo2"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maquina"
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "Inyectora"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Prensa"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Cota de control"
      Height          =   375
      Left            =   2040
      TabIndex        =   25
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Left            =   8280
      TabIndex        =   24
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label13 
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Partida"
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Compuesto"
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Moldeadas"
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Bocas"
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Matriz"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "O.T."
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Pieza"
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Lote"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmCargarLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public buscar As Boolean
Public LOTENUEVO
Public Flag
Public Function Imprime_Muestreo(ByVal moldeadas As Long, ByVal Articulo As String, ByVal lote, ByVal proceso As Boolean, ByVal cantidad As Long, ByVal cant_controles As Long, ByVal controles_intermedios As Long, ByVal cota_de_control As String, ByVal ot As String, ByVal compuestoo As String)
Dim appp As New Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim r As Excel.Range

Set wb = appp.Workbooks.Open("\\Servidor2\e\EntornoBafir\Planillas\planilla de control de proceso.xls", , True)
Set ws = wb.Worksheets(2)
ws.Cells(6, 4) = "'" & Articulo
ws.Cells(8, 4) = lote
If proceso = True Then
    ws.Cells(8, 7) = "X"
Else
    ws.Cells(6, 7) = "X"
End If
ws.Cells(10, 7) = compuestoo
ws.Cells(10, 4) = cantidad
ws.Cells(16, 3) = cota_de_control
ws.Cells(12, 9) = cant_controles
ws.Cells(12, 12) = ot
hojas_a_imprimir = CInt(((cant_controles / 20) + 0.41))
If hojas_a_imprimir = 0 Then
    hojas_a_imprimir = 1
End If
Fila = 24
moldeada = 1
Control = 1
For hoja = 1 To hojas_a_imprimir
    Fila = 24
    For renglon = 1 To 20
        If Control = cant_controles Then
            ws.Cells(Fila, 2) = moldeadas
        Else
           ''''''''''''''090225
            If Control > cant_controles Then
                Exit For
            End If
            ''''''''''''''
            If moldeada > moldeadas Then
                ws.Cells(Fila, 2) = ""
            Else
                ws.Cells(Fila, 2) = moldeada
            End If
        End If
        Fila = Fila + 1
        moldeada = moldeada + controles_intermedios
        Control = Control + 1
    Next ' renglon
    On Error Resume Next
        ws.PrintOut
        
        If Err.Number = 1004 Then
            sfsdfsdf = MsgBox("Ha ocurrido un error con la impresora. Por favor compruebe que la misma esté en condiciones de imprimir y reintente.", vbCritical + vbOKOnly, "Error")
        Else
            If Err.Number <> 0 Then
                sfsdfsdf = MsgBox("Ha ocurrido el error " & Err.Number & ". Por favor informe al administrador del sistema.", vbCritical + vbOKOnly, "Error")
            End If
        End If
        
        On Error GoTo 0
    DoEvents
Next ' Hoja
wb.Close (False)
End Function

Public Function calcula_controles_intermedios(ByVal muestreo As Long, ByVal moldeadas As Long)
controles_intermedios = muestreo - 2
If controles_intermedios <= 0 Then
    intervalo = 0
Else
    intervalo = moldeadas / (muestreo - 1)
End If
calcula_controles_intermedios = intervalo
End Function

Public Function Calcula_Muestreo(ByVal moldeadas As Long)
If Option1.Value = 1 Then  'Prensa
    muestreo = CInt(((((moldeadas + 1) ^ (1 / 2)) / 2) + 0.4))
Else ' Inyectora
    muestreo = CInt(((((moldeadas) ^ (1 / 3)) / 2) + 0.4))
End If
    Calcula_Muestreo = muestreo
End Function

Sub carga_cota()
If Flag = 0 Then
    Exit Sub
End If
codigopieza = Text1.Text
If Text1.Text = "" Then
    Exit Sub
End If
'Text4.SetFocus
codigomatriz = List1.Text
If List1.Text = "" Then
    Exit Sub
End If
codigocompuesto = Combo1.Text
If Combo1.Text = "" Then
    Exit Sub
End If
'punt = InStr(codigopieza, "-")

'If punt = 0 Then
    codigoantiguo = codigopieza
'Else
'    codigoantiguo = Left(codigopieza, punt)
'End If

Dim db As Database
Dim db1 As Database
Dim rs As Recordset
Dim rs1 As Recordset

Set db = OpenDatabase("\\Servidor2\e\entornobafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
'090223 Modificacion realizada a solicitud de Gustavo
'Set rs = db.OpenRecordset("Select * from tabla_piezas where Nro_pieza = '" & codigoantiguo & "' AND MATRIZ = '" & codigomatriz & "' AND CODIGO_COMPUESTO = '" & codigocompuesto & "'")
Set rs = db.OpenRecordset("Select * from tabla_piezas where Nro_pieza = '" & codigoantiguo & "' AND MATRIZ = '" & codigomatriz & "'")
If rs.RecordCount > 0 Then
    rs.MoveLast
End If
If rs.RecordCount = 0 Then
    asdasd = MsgBox("No se ha encontrado equivalente para la pieza en tabla de piezas. Por favor controle el artículo", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If
If rs.RecordCount > 1 Then
    asdasd = MsgBox("Existen varias piezas que se ajustan a la codificación ingresada. Por favor controle el artículo", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

Set db1 = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, False)
Set rs1 = db1.OpenRecordset("select * from cotas_temporal")
    Do Until rs1.EOF = True
        rs1.Delete
        rs1.MoveNext
    Loop
    

If rs.Fields("Medir_cota_1").Value = True Then
    rs1.AddNew
    rs1.Fields("cota") = rs.Fields("cota_1")
    rs1.Update
End If
If rs.Fields("Medir_cota_2").Value = True Then
    rs1.AddNew
    rs1.Fields("cota") = rs.Fields("cota_2")
    rs1.Update
End If
If rs.Fields("Medir_cota_3").Value = True Then
    rs1.AddNew
    rs1.Fields("cota") = rs.Fields("cota_3")
    rs1.Update
End If
If rs.Fields("Medir_cota_4").Value = True Then
    rs1.AddNew
    rs1.Fields("cota") = rs.Fields("cota_4")
    rs1.Update
End If
db1.Close
db.Close
End Sub

Function buscaPartida(ByVal comPuesto)
    Dim db As Database
    Dim rs As Recordset
    
    'Set db = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, False)
    'Set rs = db.OpenRecordset("select PARTIDA from FORMBASE WHERE N_FORMULA = '" & comPuesto & "'")
    Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
    Set rs = db.OpenRecordset("select PARTIDA from FORMBASE WHERE N_FORMULA = '" & comPuesto & "'")

    If rs.RecordCount = 0 Then
        parTida = "N/A"
    Else
        parTida = rs.Fields("Partida")
    End If
    
    buscaPartida = parTida
    db.Close
End Function

Sub ReservaLote()


    Dim db As Database
    Dim rs As Recordset
    
    Set db = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, False)
    Set rs = db.OpenRecordset("select NRO_LOTE,CDGO_PIEZA,CANT_PIEZA,COMPUESTO,FECHA,PARTIDA,OBSERVA1,NRO_MATRIZ,FRECUENCIA_CONTROL,NIVEL_DE_INSPECCION,NIVEL_DE_ACEPTACION from LOTES ORDER BY NRO_LOTE")
    rs.MoveLast
    LOTENUEVO = rs.Fields("NRO_LOTE") + 1
    rs.AddNew
    rs.Fields("NRO_LOTE") = LOTENUEVO
    rs.Fields("CDGO_PIEZA") = "0"
    rs.Fields("CANT_PIEZA") = 0
    rs.Fields("COMPUESTO") = "0"
    rs.Fields("FECHA") = 0
    rs.Fields("PARTIDA") = "0"
    rs.Fields("OBSERVA1") = "0"
    rs.Fields("NRO_MATRIZ") = "0"
    rs.Fields("FRECUENCIA_CONTROL") = 0
    rs.Fields("NIVEL_DE_INSPECCION") = 0
    rs.Fields("NIVEL_DE_ACEPTACION") = 0
    rs.Update
    rs.Close
    db.Close
    frmCargarLote.Label4.Caption = LOTENUEVO

End Sub
Sub ClearFormulario()
    Flag = 0
    Label2.Caption = Date
    Label4.Caption = LOTENUEVO
    Label14.Caption = ""
    'Cuadro_combinado11.SetFocus
    Combo1.Text = ""
    'On Error Resume Next
    'Do Until Err.Number = 0
    '    Texto6.SetFocus
    '    Err.Clear
    'Loop
    Text1.Text = ""
    'Texto39.SetFocus
    List2.Text = ""
    'Texto37.SetFocus
    Combo2.Text = ""
    'Texto8.SetFocus
    List1.Text = ""
    'Texto18.SetFocus
    Text2.Text = ""
    Text3.Text = ""
    'Texto6.SetFocus
    'On Error GoTo 0
    'Text1.SetFocus
    Combo1.Clear
    Combo2.Clear
    List1.Clear
    List2.Clear
    Label13.Caption = ""
    Flag = 1
End Sub


Private Sub Combo1_GotFocus()
'SELECT FORMBASE.N_FORMULA FROM FORMBASE WHERE FORMBASE.ESTADO = 1 OR FORMBASE.ESTADO =3;

Combo1.Clear
'Dim db1 As Database
'Dim rs1 As Recordset

'Set db1 = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, True)
'Set rs1 = db1.OpenRecordset("SELECT FORMBASE.N_FORMULA FROM FORMBASE WHERE FORMBASE.ESTADO = 1 OR FORMBASE.ESTADO =3;")

'Do Until rs1.EOF = True
'    Combo1.AddItem (rs1.Fields("N_FORMULA"))
'    rs1.MoveNext
'Loop
'db1.Close

Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("\\Servidor2\e\EntornoBafir\partidas de compuesto.mdb", False, False, ";pwd=flanflus")
Set rs = db.OpenRecordset("SELECT N_FORMULA FROM Formbase")

rs.MoveFirst
'a = rs.Fields("N_formula")

        Do While rs.EOF <> True

            rs.MoveNext
        Loop
        rs.MovePrevious
        Fila = rs.RecordCount
        rs.MoveFirst
        For contador = 1 To Fila
            b = rs.Fields("N_FORMULA").Value
        'On Error GoTo fIn
            Combo1.AddItem (rs.Fields("N_FORMULA"))
            rs.MoveNext
        Next

db.Close


End Sub

Private Sub Combo1_LostFocus()
    parTida123 = buscaPartida(Combo1.Text)
    Label14.Caption = parTida123
    carga_cota 'temporal solicitado por g.Paludi
    
End Sub

Private Sub Combo2_GotFocus()

'SELECT Cotas_Temporal.Cota FROM Cotas_Temporal;

Combo2.Clear
Dim db1 As Database
Dim rs1 As Recordset

Set db1 = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, True)
Set rs1 = db1.OpenRecordset("SELECT Cotas_Temporal.Cota FROM Cotas_Temporal;")

If rs1.RecordCount = 0 Then
    fsdfsdfsdf = MsgBox("Hubo un problema en la carga de las cotas temporales. Controle que la pieza tenga al menos un control habilitado en tabla de piezas. De lo contrario, consulte si la información ingresada es correcta o recurra al administrador del sistema", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

Do Until rs1.EOF = True
    Combo2.AddItem (rs1.Fields("cota"))
    rs1.MoveNext
Loop
db1.Close

End Sub

Private Sub Command1_Click()
    'Texto6.SetFocus
    If Text1.Text = "" Then
        pappappa = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
        Text1.SetFocus
        Exit Sub
    End If
    'Texto8.SetFocus
    If List1.Text = "" Then
        pappappa = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
        List1.SetFocus
        Exit Sub
    End If
    'Cuadro_combinado11.SetFocus
    If Combo1.Text = "" Then
        pappappa = MsgBox("Debe completar todos los campos", vbCritical + vbOKOnly, "Error")
        Combo1.SetFocus
        Exit Sub
    End If
'''''''''''''''''aca completa el registro con el lote nuevo
    Dim db As Database
    Dim rs As Recordset
    
    a = Calcula_Muestreo(CLng(Label13.Caption))
    b = calcula_controles_intermedios(a, CLng(Label13.Caption))
    
    
    
    Set db = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, False)
    Set rs = db.OpenRecordset("select OT,NRO_LOTE,CDGO_PIEZA,CANT_PIEZA,COMPUESTO,FECHA,PARTIDA,OBSERVA1,NRO_MATRIZ,FRECUENCIA_CONTROL,NIVEL_DE_INSPECCION,NIVEL_DE_ACEPTACION,MAQUINA,COTA_CONTROL1 from LOTES where NRO_LOTE = " & LOTENUEVO)
    rs.Edit
    rs.Fields("CDGO_PIEZA") = Text1.Text
    
    rs.Fields("CANT_PIEZA") = Text2.Text
    
    rs.Fields("COMPUESTO") = Combo1.Text
    rs.Fields("FECHA") = Label2.Caption
    rs.Fields("PARTIDA") = Label14.Caption
    rs.Fields("OBSERVA1") = "0"
    
    rs.Fields("NRO_MATRIZ") = List1.Text
    rs.Fields("FRECUENCIA_CONTROL") = a
    rs.Fields("NIVEL_DE_INSPECCION") = 0
    rs.Fields("NIVEL_DE_ACEPTACION") = 0
    rs.Fields("COTA_CONTROL1") = Combo2.Text
    rs.Fields("MAQUINA") = CBool(Option1.Value)
    rs.Fields("OT") = Text3.Text
    rs.Update
    rs.Close
    db.Close
    pieza = Text1.Text
    cantidadp = Text2.Text
    cota = Combo2.Text
    
    
    j = Imprime_Muestreo(CLng(Label13.Caption), pieza, LOTENUEVO, Option1.Value, CLng(cantidadp), CLng(a), CLng(b), cota, Text3.Text, Combo1.Text)
    
    LOTENUEVO = 0
    ReservaLote
    ClearFormulario
End Sub

Private Sub Command2_Click()
Flag = 0
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, False)
    Set rs = db.OpenRecordset("Select NRO_LOTE,CDGO_PIEZA,CANT_PIEZA,COMPUESTO,FECHA,PARTIDA,OBSERVA1,NRO_MATRIZ,FRECUENCIA_CONTROL,NIVEL_DE_INSPECCION,NIVEL_DE_ACEPTACION from LOTES where NRO_LOTE = " & LOTENUEVO)
    If rs.RecordCount <> 0 Then
        rs.Delete
    End If
    rs.Close
    db.Close
Form1.Enabled = True
Unload Me
End Sub



Private Sub Form_Load()
    If buscar = True Then
    
    Else
        ReservaLote
        ClearFormulario
    End If
End Sub

Private Sub List1_GotFocus()
List1.Clear
Dim db1 As Database
Dim rs1 As Recordset

Set db1 = OpenDatabase("\\Servidor2\e\entornobafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db1.OpenRecordset("SELECT tabla_piezas.Matriz FROM tabla_piezas Where tabla_piezas.nro_pieza = '" & Text1.Text & "'")

If rs1.RecordCount = 0 Then
    sfdsdfgdfg = MsgBox("No existen datos para su búsqueda. Por favor consulte que sus datos sean correctos", vbCritical + vbOKOnly, "Error")
    Exit Sub
End If

Do Until rs1.EOF = True
    List1.AddItem (rs1.Fields("matriz"))
    rs1.MoveNext
Loop
db1.Close
End Sub



Private Sub List2_GotFocus()
'SELECT tabla_piezas.bocas FROM tabla_piezas Where tabla_piezas.nro_pieza = texto6.value AND tabla_piezas.matriz = texto8.value;

List2.Clear
Dim db1 As Database
Dim rs1 As Recordset
Set db1 = OpenDatabase("\\Servidor2\e\entornobafir\partidas de compuesto.mdb", False, True, ";pwd=flanflus")
Set rs1 = db1.OpenRecordset("SELECT tabla_piezas.bocas FROM tabla_piezas Where tabla_piezas.nro_pieza = '" & Text1.Text & "' AND tabla_piezas.matriz = '" & List1.Text & "'")

If rs1.RecordCount = 0 Then
    sdfsdfsdfsdf = MsgBox("No existen datos para su búsqueda. Por favor controle que la información sea correcta", vbCritical + vbOKOnly, "Error")
End If

Do Until rs1.EOF = True
    List2.AddItem (rs1.Fields("bocas"))
    rs1.MoveNext
Loop
db1.Close


End Sub

Private Sub List2_LostFocus()
If Flag = 1 Then
If List2.Text = "" Then
    asdasd = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    List2.SetFocus
    Exit Sub
End If
If IsNumeric(List2.Text) = False Then
    asdasd = MsgBox("Debe completar el campo con valores coherentes", vbCritical + vbOKOnly, "Error")
    List2.SetFocus
    Exit Sub
End If
Label13.Caption = CLng(Text2.Text / List2.Text)
End If
End Sub

Private Sub Text1_LostFocus()
Dim asdasd
'If Flag = 1 Then
    'If Text1.Text = "" Then
    '    asdasd = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    '    Text1.SetFocus
    'End If
'End If
Dim db1 As Database
Dim rs1 As Recordset

Set db1 = OpenDatabase("\\Servidor2\e\produccion\lotes\loteproducción.mdb", False, False)
Set rs1 = db1.OpenRecordset("select * from cotas_temporal")
    Do Until rs1.EOF = True
        rs1.Delete
        rs1.MoveNext
    Loop
db1.Close
End Sub
Private Sub Text2_LostFocus()
Dim asdasd
If Flag = 1 Then
If Text2.Text = "" Then
    asdasd = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Texto2.SetFocus
End If
If IsNumeric(Text2.Text) = False Then
    asdasd = MsgBox("Debe completar el campo con valores coherentes", vbCritical + vbOKOnly, "Error")
    Texto2.SetFocus
End If
End If
End Sub
Private Sub Text3_LostFocus()
If Flag = 1 Then
If Text3.Text = "" Then
    asdasd = MsgBox("Debe completar el campo", vbCritical + vbOKOnly, "Error")
    Text3.SetFocus
End If
If IsNumeric(Text3.Text) = False Then
    asdasd = MsgBox("Debe completar el campo con valores coherentes", vbCritical + vbOKOnly, "Error")
    Text3.SetFocus
End If


    
    
    codigopieza = Text1.Text
    
    punt = InStr(codigopieza, "-")
    
    If punt = 0 Then
        codigoantiguo = codigopieza
    Else
        codigoantiguo = Left(codigopieza, punt)
    End If
    
       
End If


End Sub
